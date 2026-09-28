"""
Leitura e gravação dos dados da empresa em planilhas CSV.

É por aqui que a IA se liga aos seus dados reais: exporte do ERP (ou preencha
à mão) os arquivos abaixo em uma pasta e rode `python -m pcp_ia planejar`.

Cadastros:
  itens.csv               codigo, descricao, tipo (fabricado/comprado), lead_time,
                          lote_minimo, lote_multiplo, estoque_seguranca, refugo,
                          centro, centro_alternativo, horas_por_unidade, fornecedor, custo
  estrutura.csv           pai, filho, quantidade                (lista de materiais / BOM)
  centros.csv             centro, horas_semana                  (capacidade nominal)
  premissas.json          regras de planejamento (opcional)

Situação atual:
  estoque.csv             item, quantidade
  demanda_historica.csv   semana, item, quantidade              (vendas/consumo por semana)
  pedidos_carteira.csv    item, semana, quantidade              (pedidos firmes; opcional)
  ordens_abertas.csv      item, quantidade, semana_entrega      (OPs e pedidos de compra em aberto)
  pendencias.csv          item, quantidade                      (atrasado com o cliente; opcional)

Realizado (é daqui que a IA aprende; vá acrescentando linhas a cada semana):
  realizado_compras.csv   item, semana_emissao, semana_recebimento
  realizado_producao.csv  semana, item, centro, quantidade, quantidade_boa
  realizado_eficiencia.csv semana, centro, eficiencia

As semanas são números inteiros e crescentes (ex.: 1, 2, 3... ou 202601, 202602...
desde que não pulem de ano; prefira uma numeração contínua).

A pasta `modelo/` guarda o que a IA aprendeu entre um ciclo e outro: os pesos
da rede neural e as previsões feitas (para medir o erro quando a demanda real chegar).
"""

import csv
import json
import os

import numpy as np

from .mrp import Item
from .planejador import PlanejadorPCP
from .rede_neural import RedeNeural

CAMPOS_ITEM = ["codigo", "descricao", "tipo", "lead_time", "lote_minimo", "lote_multiplo",
               "estoque_seguranca", "refugo", "centro", "centro_alternativo",
               "horas_por_unidade", "fornecedor", "custo"]


# ---------------------------------------------------------------- utilitários CSV
def ler_csv(caminho, obrigatorio=True):
    if not os.path.exists(caminho):
        if obrigatorio:
            raise FileNotFoundError(f"Arquivo obrigatório não encontrado: {caminho}")
        return []
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        amostra = f.read(2048)
        f.seek(0)
        separador = ";" if amostra.count(";") > amostra.count(",") else ","
        return [
            {k.strip(): (v or "").strip() for k, v in linha.items() if k}
            for linha in csv.DictReader(f, delimiter=separador)
        ]


def escrever_csv(caminho, linhas, campos=None):
    campos = campos or (list(linhas[0].keys()) if linhas else [])
    with open(caminho, "w", encoding="utf-8", newline="") as f:
        escritor = csv.DictWriter(f, fieldnames=campos)
        escritor.writeheader()
        escritor.writerows(linhas)


def numero(valor, padrao=0.0):
    if valor in (None, ""):
        return padrao
    return float(str(valor).replace(",", "."))


# ---------------------------------------------------------------- leitura
def carregar_pasta(pasta):
    j = lambda nome: os.path.join(pasta, nome)  # noqa: E731
    itens = {}
    for l in ler_csv(j("itens.csv")):
        itens[l["codigo"]] = Item(
            codigo=l["codigo"], descricao=l.get("descricao", ""), tipo=l["tipo"].lower(),
            lead_time=int(numero(l.get("lead_time"))), lote_minimo=numero(l.get("lote_minimo")),
            lote_multiplo=numero(l.get("lote_multiplo"), 1) or 1,
            estoque_seguranca=numero(l.get("estoque_seguranca")), refugo=numero(l.get("refugo")),
            centro=l.get("centro", ""), centro_alternativo=l.get("centro_alternativo", ""),
            horas_por_unidade=numero(l.get("horas_por_unidade")),
            fornecedor=l.get("fornecedor", ""), custo=numero(l.get("custo")),
        )
    estrutura = [(l["pai"], l["filho"], numero(l["quantidade"])) for l in ler_csv(j("estrutura.csv"))]
    capacidade = {l["centro"]: numero(l["horas_semana"]) for l in ler_csv(j("centros.csv"))}

    premissas = {}
    if os.path.exists(j("premissas.json")):
        with open(j("premissas.json"), encoding="utf-8") as f:
            premissas = json.load(f)

    _validar(itens, estrutura, capacidade)
    return {
        "itens": itens,
        "estrutura": estrutura,
        "capacidade": capacidade,
        "premissas": premissas,
        "estoque": {l["item"]: numero(l["quantidade"]) for l in ler_csv(j("estoque.csv"))},
        "demanda": ler_csv(j("demanda_historica.csv")),
        "carteira": ler_csv(j("pedidos_carteira.csv"), obrigatorio=False),
        "abertas": ler_csv(j("ordens_abertas.csv"), obrigatorio=False),
        "pendencias": {l["item"]: numero(l["quantidade"])
                       for l in ler_csv(j("pendencias.csv"), obrigatorio=False)},
        "compras": ler_csv(j("realizado_compras.csv"), obrigatorio=False),
        "producao": ler_csv(j("realizado_producao.csv"), obrigatorio=False),
        "eficiencia": ler_csv(j("realizado_eficiencia.csv"), obrigatorio=False),
    }


def _validar(itens, estrutura, capacidade):
    erros = []
    for pai, filho, _ in estrutura:
        for codigo in (pai, filho):
            if codigo not in itens:
                erros.append(f"estrutura.csv cita o item '{codigo}', que não está em itens.csv")
    for item in itens.values():
        if item.tipo not in ("fabricado", "comprado"):
            erros.append(f"Item {item.codigo}: tipo deve ser 'fabricado' ou 'comprado'")
        if item.tipo == "fabricado":
            for centro in (item.centro, item.centro_alternativo):
                if centro and centro not in capacidade:
                    erros.append(f"Item {item.codigo}: centro '{centro}' não está em centros.csv")
            if not item.centro:
                erros.append(f"Item {item.codigo}: item fabricado precisa de um centro")
    if erros:
        raise ValueError("Problemas nos dados:\n  - " + "\n  - ".join(erros))


# ---------------------------------------------------------------- ciclo real
def planejar_pasta(pasta):
    """Roda um ciclo de planejamento com os dados da pasta. Devolve (planejador, plano)."""
    d = carregar_pasta(pasta)
    planejador = PlanejadorPCP(d["itens"], d["estrutura"], d["capacidade"], d["premissas"])
    horizonte = planejador.premissas["horizonte_semanas"]
    memoria = planejador.memoria

    # Série semanal de demanda de cada produto acabado (semanas sem venda = 0)
    semanas = [int(numero(l["semana"])) for l in d["demanda"]]
    if not semanas:
        raise ValueError("demanda_historica.csv está vazio")
    primeira, ultima = min(semanas), max(semanas)
    semana_atual = ultima + 1
    historico = {p: np.zeros(ultima - primeira + 1) for p in planejador.acabados}
    for l in d["demanda"]:
        if l["item"] in historico:
            historico[l["item"]][int(numero(l["semana"])) - primeira] += numero(l["quantidade"])

    # Aprendizado a partir do realizado
    for l in d["compras"]:
        if l.get("semana_recebimento"):
            memoria.registrar_lead_time(
                l["item"], int(numero(l["semana_recebimento"])) - int(numero(l["semana_emissao"])))
    for l in sorted(d["producao"], key=lambda l: numero(l.get("semana"))):
        memoria.registrar_producao(l["item"], numero(l["quantidade"]), numero(l["quantidade_boa"]))
    for l in sorted(d["eficiencia"], key=lambda l: numero(l.get("semana"))):
        memoria.registrar_eficiencia(l["centro"], numero(l["eficiencia"]))

    pasta_modelo = os.path.join(pasta, "modelo")
    os.makedirs(pasta_modelo, exist_ok=True)
    arquivo_previsoes = os.path.join(pasta_modelo, "previsoes.csv")
    previsoes_feitas = {}
    for l in ler_csv(arquivo_previsoes, obrigatorio=False):
        previsoes_feitas[(int(numero(l["semana"])), l["item"])] = numero(l["previsto"])
    for (semana, item), previsto in sorted(previsoes_feitas.items()):
        if item in historico and primeira <= semana <= ultima:
            memoria.registrar_previsao(item, previsto, historico[item][semana - primeira])

    # A rede neural continua de onde parou no ciclo anterior
    arquivo_rede = os.path.join(pasta_modelo, "rede_neural.npz")
    if os.path.exists(arquivo_rede):
        rede = RedeNeural.carregar(arquivo_rede, seed=planejador.previsor.seed)
        # Se o mix de produtos mudou, a rede antiga não serve: treina uma nova
        if rede.n_entradas == planejador.previsor.n_entradas:
            planejador.previsor.rede = rede

    recebimentos = {c: [0.0] * horizonte for c in d["itens"]}
    for l in d["abertas"]:
        t = max(0, int(numero(l["semana_entrega"])) - semana_atual)
        if t < horizonte:
            recebimentos[l["item"]][t] += numero(l["quantidade"])
    carteira = {}
    for l in d["carteira"]:
        t = int(numero(l["semana"])) - semana_atual
        if 0 <= t < horizonte:
            carteira.setdefault(l["item"], np.zeros(horizonte))[t] += numero(l["quantidade"])
        elif t < 0:  # pedido vencido e não entregue = pendência
            d["pendencias"][l["item"]] = d["pendencias"].get(l["item"], 0) + numero(l["quantidade"])

    plano = planejador.planejar(
        semana_atual, {p: list(s) for p, s in historico.items()}, d["estoque"], recebimentos,
        pedidos_firmes=carteira, pendencias=d["pendencias"])

    # Guarda o que foi aprendido para o próximo ciclo
    if planejador.previsor.rede is not None:
        planejador.previsor.rede.salvar(arquivo_rede)
    for p in planejador.acabados:
        previsoes_feitas[(semana_atual, p)] = round(float(plano.previsao[p][0]), 2)
    escrever_csv(arquivo_previsoes, [
        {"semana": s, "item": i, "previsto": v} for (s, i), v in sorted(previsoes_feitas.items())
    ], ["semana", "item", "previsto"])
    memoria.salvar(os.path.join(pasta_modelo, "memoria_aprendizado.json"))
    return planejador, plano


# ---------------------------------------------------------------- exemplo
def gerar_pasta_exemplo(pasta, ciclos=26, seed=42):
    """Cria uma pasta com os dados da empresa fictícia, no formato dos CSVs reais."""
    from .simulacao import empresa_exemplo, simular

    os.makedirs(pasta, exist_ok=True)
    r = simular(ciclos=ciclos, aprender=True, seed=seed)
    itens, estrutura, capacidade = empresa_exemplo()
    j = lambda nome: os.path.join(pasta, nome)  # noqa: E731

    escrever_csv(j("itens.csv"), [
        {campo: getattr(item, campo) for campo in CAMPOS_ITEM} for item in itens.values()
    ], CAMPOS_ITEM)
    escrever_csv(j("estrutura.csv"), [
        {"pai": p, "filho": f, "quantidade": q} for p, f, q in estrutura])
    escrever_csv(j("centros.csv"), [
        {"centro": c, "horas_semana": h} for c, h in capacidade.items()])
    escrever_csv(j("estoque.csv"), [
        {"item": c, "quantidade": round(q, 2)} for c, q in r.estoque.items()])
    escrever_csv(j("demanda_historica.csv"), [
        {"semana": s, "item": p, "quantidade": q}
        for p, serie in r.historico.items() for s, q in enumerate(serie)])
    escrever_csv(j("ordens_abertas.csv"), [
        {"item": o["item"], "quantidade": round(o["quantidade_esperada"], 2),
         "semana_entrega": o["prometida"]} for o in r.ordens_abertas
    ], ["item", "quantidade", "semana_entrega"])
    escrever_csv(j("pendencias.csv"), [
        {"item": p, "quantidade": round(q, 2)} for p, q in r.pendencias.items() if q > 0
    ], ["item", "quantidade"])
    escrever_csv(j("pedidos_carteira.csv"), [
        {"item": "MES", "semana": r.semana_final + 1, "quantidade": 80},
    ])
    escrever_csv(j("realizado_compras.csv"), r.realizado_compras,
                 ["item", "semana_emissao", "semana_recebimento"])
    escrever_csv(j("realizado_producao.csv"), r.realizado_producao,
                 ["semana", "item", "centro", "quantidade", "quantidade_boa"])
    escrever_csv(j("realizado_eficiencia.csv"), r.realizado_eficiencia,
                 ["semana", "centro", "eficiencia"])
    with open(j("premissas.json"), "w", encoding="utf-8") as f:
        json.dump(r.planejador.premissas, f, ensure_ascii=False, indent=2)

    os.makedirs(j("modelo"), exist_ok=True)
    escrever_csv(os.path.join(pasta, "modelo", "previsoes.csv"), r.previsoes,
                 ["semana", "item", "previsto"])
    r.planejador.previsor.rede.salvar(os.path.join(pasta, "modelo", "rede_neural.npz"))
    return pasta
