"""
Empresa fictícia para testar a IA antes de ligar nos dados reais.

"Móveis Exemplo" fabrica cadeiras, mesas e bancos. Como em qualquer fábrica
real, o cadastro do ERP está errado em vários pontos, e só a "realidade"
(escondida do planejador) sabe os valores verdadeiros:

  * a madeira leva 3 semanas para chegar, mas o cadastro diz 2
  * a montagem rende 82% da capacidade nominal, mas o cadastro diz 100%
  * a produção perde de 2% a 6% por refugo, mas o cadastro diz 0%
  * a demanda tem sazonalidade: pico de mesas no fim de ano e de bancos no verão

A simulação roda semana a semana. Em cada semana o planejador faz o plano, a
fábrica executa e o resultado real volta para o aprendizado.
"""

import math
from dataclasses import dataclass, field

import numpy as np

from .mrp import Item
from .planejador import PlanejadorPCP
from .previsao import SEMANAS_ANO


def empresa_exemplo():
    itens = {
        "CAD": Item("CAD", "Cadeira", "fabricado", 1, lote_multiplo=10, estoque_seguranca=20,
                    centro="MONTAGEM", centro_alternativo="MONTAGEM2", horas_por_unidade=0.5, custo=180),
        "MES": Item("MES", "Mesa", "fabricado", 1, lote_multiplo=5, estoque_seguranca=8,
                    centro="MONTAGEM", centro_alternativo="MONTAGEM2", horas_por_unidade=1.2, custo=520),
        "BAN": Item("BAN", "Banco", "fabricado", 1, lote_multiplo=10, estoque_seguranca=10,
                    centro="MONTAGEM", centro_alternativo="MONTAGEM2", horas_por_unidade=0.3, custo=90),
        "ASS": Item("ASS", "Assento estofado", "fabricado", 1, lote_multiplo=20,
                    centro="ESTOFARIA", horas_por_unidade=0.25, custo=45),
        "MAD": Item("MAD", "Madeira (m)", "comprado", 2, lote_multiplo=50, fornecedor="Madeireira Sul", custo=18),
        "PAR": Item("PAR", "Parafuso", "comprado", 1, lote_minimo=1000, lote_multiplo=500,
                    fornecedor="Fixa Tudo", custo=0.15),
        "ESP": Item("ESP", "Espuma", "comprado", 2, lote_multiplo=20, fornecedor="Espumas SP", custo=12),
        "TEC": Item("TEC", "Tecido (m)", "comprado", 3, lote_multiplo=25, fornecedor="Têxtil Norte", custo=22),
        "VER": Item("VER", "Verniz (L)", "comprado", 1, lote_multiplo=5, fornecedor="Química Ltda", custo=35),
    }
    estrutura = [
        ("CAD", "ASS", 1), ("CAD", "MAD", 2.0), ("CAD", "PAR", 8), ("CAD", "VER", 0.1),
        ("MES", "MAD", 6.0), ("MES", "PAR", 12), ("MES", "VER", 0.4),
        ("BAN", "ASS", 1), ("BAN", "MAD", 1.5), ("BAN", "PAR", 4), ("BAN", "VER", 0.05),
        ("ASS", "ESP", 1), ("ASS", "TEC", 0.5),
    ]
    capacidade_horas = {"MONTAGEM": 180, "MONTAGEM2": 40, "ESTOFARIA": 80}
    return itens, estrutura, capacidade_horas


class Realidade:
    """O que realmente acontece na fábrica (o planejador não enxerga isto)."""

    lead_time = {"MAD": (3, 0.7), "PAR": (1, 0.3), "ESP": (3, 0.8), "TEC": (3, 0.5), "VER": (2, 0.4)}
    refugo = {"CAD": 0.03, "MES": 0.05, "BAN": 0.02, "ASS": 0.06}
    eficiencia = {"MONTAGEM": 0.82, "MONTAGEM2": 0.90, "ESTOFARIA": 0.88}
    demanda_base = {"CAD": 120, "MES": 40, "BAN": 60}

    def __init__(self, seed=42):
        self.rng = np.random.default_rng(seed)

    def demanda(self, produto, semana):
        s = semana % SEMANAS_ANO
        tendencia = 1 + 0.0005 * semana
        if produto == "CAD":
            sazonal = 1 + 0.25 * math.sin(2 * math.pi * s / SEMANAS_ANO)
        elif produto == "MES":
            sazonal = 1.7 if 44 <= s <= 50 else 1 + 0.1 * math.cos(2 * math.pi * s / SEMANAS_ANO)
        else:  # BAN: pico no verão (começo do ano)
            sazonal = 1 + 0.4 * math.cos(2 * math.pi * s / SEMANAS_ANO)
        media = self.demanda_base[produto] * tendencia * sazonal
        return max(0, int(round(self.rng.normal(media, 0.06 * media))))

    def sortear_lead_time(self, item):
        media, desvio = self.lead_time[item]
        return max(1, int(round(self.rng.normal(media, desvio))))

    def sortear_eficiencia(self, centro):
        return float(np.clip(self.rng.normal(self.eficiencia[centro], 0.03), 0.5, 1.0))

    def produzir(self, item, quantidade):
        return int(self.rng.binomial(int(quantidade), 1 - self.refugo[item]))


@dataclass
class ResultadoSimulacao:
    metricas: list = field(default_factory=list)
    planejador: PlanejadorPCP = None
    ultimo_plano: object = None
    # Estado final e registros do "chão de fábrica" (usados para gerar os CSVs de exemplo)
    semana_final: int = 0
    historico: dict = None
    estoque: dict = None
    pendencias: dict = None
    ordens_abertas: list = None
    realizado_compras: list = field(default_factory=list)
    realizado_producao: list = field(default_factory=list)
    realizado_eficiencia: list = field(default_factory=list)
    previsoes: list = field(default_factory=list)


def simular(ciclos=52, aprender=True, seed=42, semanas_historico=104, premissas=None,
            ao_final_do_ciclo=None):
    itens, estrutura, capacidade = empresa_exemplo()
    realidade = Realidade(seed)
    planejador = PlanejadorPCP(itens, estrutura, capacidade, premissas, aprender, seed=seed)
    horizonte = planejador.premissas["horizonte_semanas"]
    acabados = planejador.acabados

    historico = {p: [realidade.demanda(p, t) for t in range(semanas_historico)] for p in acabados}
    estoque = {c: 0.0 for c in itens}
    for p in acabados:
        estoque[p] = 2.0 * np.mean(historico[p][-8:])
    estoque.update({"ASS": 400, "MAD": 2500, "PAR": 8000, "ESP": 700, "TEC": 350, "VER": 100})
    pendencias = {p: 0.0 for p in acabados}
    abertas = []  # ordens em andamento
    resultado = ResultadoSimulacao(planejador=planejador)

    for ciclo in range(ciclos):
        semana = semanas_historico + ciclo

        # 1. Chegam as ordens previstas para esta semana
        for ordem in [o for o in abertas if o["chegada"] <= semana]:
            estoque[ordem["item"]] += ordem["quantidade_real"]
            if ordem["tipo"] == "comprar":
                planejador.registrar_recebimento_compra(ordem["item"], ordem["chegada"] - ordem["emissao"])
                resultado.realizado_compras.append(
                    {"item": ordem["item"], "semana_emissao": ordem["emissao"], "semana_recebimento": ordem["chegada"]})
            abertas.remove(ordem)

        # 2. Planejamento: o planejador só enxerga as datas prometidas
        recebimentos = {c: [0.0] * horizonte for c in itens}
        for ordem in abertas:
            t = min(max(0, ordem["prometida"] - semana), horizonte - 1)
            recebimentos[ordem["item"]][t] += ordem["quantidade_esperada"]
        plano = planejador.planejar(semana, historico, estoque, recebimentos, pendencias=pendencias)

        # 3. Execução: libera as ordens desta semana
        horas_livres = {}
        for centro, horas in capacidade.items():
            eficiencia = realidade.sortear_eficiencia(centro)
            horas_livres[centro] = horas * eficiencia
            planejador.registrar_eficiencia(centro, eficiencia)
            resultado.realizado_eficiencia.append(
                {"semana": semana, "centro": centro, "eficiencia": round(eficiencia, 4)})

        faltas_material = 0
        liberar = [o for o in plano.ordens if o.semana_liberacao == 0]
        # Componentes (níveis mais baixos) primeiro: assento antes de cadeira
        liberar.sort(key=lambda o: (o.tipo != "comprar", o.item in acabados, o.semana_necessidade))
        for ordem in liberar:
            item = itens[ordem.item]
            lead_planejado = plano.parametros[ordem.item]["lead_time"]
            if ordem.tipo == "comprar":
                abertas.append({
                    "item": ordem.item, "tipo": "comprar", "emissao": semana,
                    "chegada": semana + realidade.sortear_lead_time(ordem.item),
                    "prometida": semana + lead_planejado,
                    "quantidade_esperada": ordem.quantidade, "quantidade_real": ordem.quantidade,
                })
                continue

            componentes = [(f, q) for p, f, q in estrutura if p == ordem.item]
            limite_material = min(estoque[f] / q for f, q in componentes)
            limite_capacidade = horas_livres[ordem.centro] / item.horas_por_unidade
            quantidade = math.floor(min(ordem.quantidade, limite_material, limite_capacidade))
            if limite_material < ordem.quantidade:
                faltas_material += 1
            if quantidade <= 0:
                continue
            for filho, q in componentes:
                estoque[filho] -= quantidade * q
            horas_livres[ordem.centro] -= quantidade * item.horas_por_unidade
            boas = realidade.produzir(ordem.item, quantidade)
            planejador.registrar_producao(ordem.item, quantidade, boas)
            resultado.realizado_producao.append(
                {"semana": semana, "item": ordem.item, "centro": ordem.centro,
                 "quantidade": quantidade, "quantidade_boa": boas})
            refugo_planejado = plano.parametros[ordem.item]["refugo"]
            abertas.append({
                "item": ordem.item, "tipo": "produzir", "emissao": semana,
                "chegada": semana + item.lead_time, "prometida": semana + item.lead_time,
                "quantidade_esperada": quantidade * (1 - refugo_planejado), "quantidade_real": boas,
            })

        # 4. A demanda real acontece
        demanda_real = {p: realidade.demanda(p, semana) for p in acabados}
        demandado = atendido_no_prazo = erro_abs = 0.0
        for p in acabados:
            atrasado_atendido = min(estoque[p], pendencias[p])
            estoque[p] -= atrasado_atendido
            pendencias[p] -= atrasado_atendido
            no_prazo = min(estoque[p], demanda_real[p])
            estoque[p] -= no_prazo
            pendencias[p] += demanda_real[p] - no_prazo
            demandado += demanda_real[p]
            atendido_no_prazo += no_prazo
            erro_abs += abs(plano.previsao[p][0] - demanda_real[p])
            historico[p].append(demanda_real[p])
        planejador.registrar_demanda_real(plano, demanda_real)
        resultado.previsoes.extend(
            {"semana": semana, "item": p, "previsto": round(float(plano.previsao[p][0]), 2)}
            for p in acabados)

        metricas = {
            "ciclo": ciclo + 1,
            "semana": semana,
            "erro_previsao_pct": 100 * erro_abs / max(demandado, 1),
            "nivel_servico_pct": 100 * atendido_no_prazo / max(demandado, 1),
            "valor_estoque": sum(estoque[c] * itens[c].custo for c in itens),
            "pendencias_unid": sum(pendencias.values()),
            "faltas_material": faltas_material,
        }
        resultado.metricas.append(metricas)
        if ao_final_do_ciclo:
            ao_final_do_ciclo(metricas, plano)
        resultado.ultimo_plano = plano

    resultado.semana_final = semanas_historico + ciclos
    resultado.historico = historico
    resultado.estoque = estoque
    resultado.pendencias = pendencias
    resultado.ordens_abertas = abertas
    return resultado
