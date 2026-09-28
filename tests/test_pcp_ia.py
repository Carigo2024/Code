import os

import numpy as np
import pytest

from pcp_ia.aprendizado import MemoriaAprendizado
from pcp_ia.capacidade import alocar_capacidade
from pcp_ia.dados import gerar_pasta_exemplo, ler_csv, planejar_pasta
from pcp_ia.mrp import Item, aplicar_lote, calcular_mrp
from pcp_ia.previsao import PrevisorDemanda, previsao_media_movel
from pcp_ia.rede_neural import RedeNeural
from pcp_ia.simulacao import Realidade, simular


# ------------------------------------------------------------ rede neural
def test_rede_neural_aprende_funcao_nao_linear(tmp_path):
    rng = np.random.default_rng(0)
    X = rng.uniform(-2, 2, size=(400, 2))
    y = np.sin(X[:, 0]) + X[:, 1] ** 2
    rede = RedeNeural(2, (32, 16), taxa_aprendizado=5e-3)
    erros = rede.treinar(X, y, epocas=150)
    assert erros[-1] < erros[0] * 0.1

    caminho = tmp_path / "rede.npz"
    rede.salvar(caminho)
    copia = RedeNeural.carregar(caminho)
    np.testing.assert_allclose(copia.prever(X[:5]), rede.prever(X[:5]))


# ------------------------------------------------------------ MRP
def _itens_simples():
    return {
        "PA": Item("PA", "Produto", "fabricado", 1, lote_multiplo=10, centro="C1",
                   horas_por_unidade=1.0),
        "MP": Item("MP", "Matéria-prima", "comprado", 2, lote_minimo=100),
    }


def _parametros(itens, **extra):
    base = {c: {"lead_time": i.lead_time, "refugo": 0.0, "estoque_seguranca": 0.0}
            for c, i in itens.items()}
    for codigo, valores in extra.items():
        base[codigo].update(valores)
    return base


def test_aplicar_lote():
    item = Item("X", "", "comprado", 1, lote_minimo=100, lote_multiplo=25)
    assert aplicar_lote(0, item) == 0
    assert aplicar_lote(30, item) == 100
    assert aplicar_lote(110, item) == 125


def test_mrp_explode_estrutura_e_desloca_lead_time():
    itens = _itens_simples()
    estrutura = [("PA", "MP", 2)]
    ordens, tabela = calcular_mrp(
        itens, estrutura, {"PA": [0, 0, 0, 25, 0, 0]}, {"PA": 0, "MP": 0},
        {}, _parametros(itens), horizonte=6)

    pa = [o for o in ordens if o.item == "PA"]
    assert len(pa) == 1
    assert pa[0].quantidade == 30               # 25 arredondado ao múltiplo de 10
    assert pa[0].semana_necessidade == 3
    assert pa[0].semana_liberacao == 2          # lead time de 1 semana

    mp = [o for o in ordens if o.item == "MP"]
    assert mp[0].quantidade == 100              # 60 necessários, lote mínimo 100
    assert mp[0].semana_necessidade == 2        # quando a OP do pai é liberada
    assert mp[0].semana_liberacao == 0          # lead time de 2 semanas
    assert tabela["MP"][2]["necessidade_bruta"] == 60


def test_mrp_compensa_refugo_e_sinaliza_atraso():
    itens = _itens_simples()
    ordens, _ = calcular_mrp(
        itens, [("PA", "MP", 1)], {"PA": [90]}, {}, {},
        _parametros(itens, PA={"refugo": 0.1}), horizonte=1)
    pa = next(o for o in ordens if o.item == "PA")
    assert pa.quantidade == 100                 # 90 / (1 - 10%) = 100
    assert any("ATRASADA" in a for a in pa.alertas)


def test_mrp_desconta_estoque_e_recebimentos():
    itens = _itens_simples()
    ordens, _ = calcular_mrp(
        itens, [], {"PA": [50, 50, 50]}, {"PA": 60}, {"PA": [0, 40, 0]},
        _parametros(itens), horizonte=3)
    # 60 + 40 cobrem as duas primeiras semanas; falta 50 na terceira
    assert [(o.semana_necessidade, o.quantidade) for o in ordens] == [(2, 50)]


# ------------------------------------------------------------ capacidade
def test_capacidade_divide_ordem_para_centro_alternativo():
    itens = {"PA": Item("PA", "", "fabricado", 1, centro="C1", centro_alternativo="C2",
                        horas_por_unidade=1.0)}
    ordens, _ = calcular_mrp(itens, [], {"PA": [0, 150]}, {}, {}, _parametros(itens), 2)
    carga = alocar_capacidade(ordens, itens, {"C1": 125, "C2": 100}, {"C1": 0.8, "C2": 1.0}, 2)
    assert sorted((o.centro, o.quantidade) for o in ordens) == [("C1", 100), ("C2", 50)]
    assert carga["C1"][0] == 100 and carga["C2"][0] == 50


def test_capacidade_alerta_sobrecarga():
    itens = {"PA": Item("PA", "", "fabricado", 1, centro="C1", horas_por_unidade=2.0)}
    ordens, _ = calcular_mrp(itens, [], {"PA": [0, 100]}, {}, {}, _parametros(itens), 2)
    alocar_capacidade(ordens, itens, {"C1": 100}, {"C1": 1.0}, 2)
    assert any("SOBRECARGA" in a for a in ordens[0].alertas)


# ------------------------------------------------------------ aprendizado
def test_memoria_aprende_lead_time_refugo_e_eficiencia():
    m = MemoriaAprendizado(amostras_minimas=3)
    assert m.lead_time_planejamento("MP", cadastro=2) == 2   # sem dados: cadastro
    for lt in [3, 3, 4, 3, 3, 5, 3, 4]:
        m.registrar_lead_time("MP", lt)
    assert m.lead_time_planejamento("MP", cadastro=2, percentil=90) >= 4

    for _ in range(30):
        m.registrar_producao("PA", 100, 95)
        m.registrar_eficiencia("C1", 0.8)
    assert m.taxa_refugo("PA") == pytest.approx(0.05)
    assert m.eficiencia_centro("C1") == pytest.approx(0.8)

    for erro in [10, -10, 12, -8, 9, -11]:
        m.registrar_previsao("PA", 100, 100 + erro)
    assert m.estoque_seguranca("PA", lead_time=1, nivel_servico=0.95) > 0


def test_memoria_salva_e_carrega(tmp_path):
    m = MemoriaAprendizado()
    m.registrar_lead_time("MP", 3)
    m.registrar_producao("PA", 100, 90)
    caminho = tmp_path / "memoria.json"
    m.salvar(caminho)
    copia = MemoriaAprendizado.carregar(caminho)
    assert copia.lead_times == {"MP": [3]}
    assert copia.taxa_refugo("PA") == pytest.approx(0.1)


# ------------------------------------------------------------ previsão
def test_rede_neural_preve_melhor_que_media_movel():
    realidade = Realidade(seed=7)
    produtos = ["CAD", "MES", "BAN"]
    historico = {p: [realidade.demanda(p, t) for t in range(104)] for p in produtos}
    previsor = PrevisorDemanda(produtos, seed=0)
    erro_rede = erro_media = 0.0
    for semana in range(104, 130):
        previsor.treinar(historico)
        previsao = previsor.prever(historico, 1)
        for p in produtos:
            real = realidade.demanda(p, semana)
            erro_rede += abs(previsao[p][0] - real)
            erro_media += abs(previsao_media_movel(historico[p], 1)[0] - real)
            historico[p].append(real)
    assert erro_rede < erro_media


# ------------------------------------------------------------ ciclo completo
def test_aprendizado_melhora_o_nivel_de_servico():
    sem = simular(ciclos=40, aprender=False)
    com = simular(ciclos=40, aprender=True)

    def servico(res):
        return np.mean([m["nivel_servico_pct"] for m in res.metricas[20:]])

    assert servico(com) > servico(sem) + 30
    assert servico(com) > 85
    # Os parâmetros aprendidos devem se aproximar da realidade
    params = com.planejador.parametros()
    assert params["MAD"]["lead_time"] >= Realidade.lead_time["MAD"][0]
    assert com.planejador.eficiencias()["MONTAGEM"] == pytest.approx(0.82, abs=0.05)


def test_planejar_com_pasta_de_dados(tmp_path):
    pasta = gerar_pasta_exemplo(str(tmp_path / "dados"), ciclos=8)
    planejador, plano = planejar_pasta(pasta)
    assert plano.usou_rede_neural
    assert plano.ordens
    assert os.path.exists(os.path.join(pasta, "modelo", "memoria_aprendizado.json"))

    # Segundo ciclo: a rede continua de onde parou e a previsão fica registrada
    antes = len(ler_csv(os.path.join(pasta, "modelo", "previsoes.csv")))
    planejar_pasta(pasta)
    depois = len(ler_csv(os.path.join(pasta, "modelo", "previsoes.csv")))
    assert depois == antes  # mesma semana: a previsão é substituída, não duplicada
