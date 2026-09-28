"""
O planejador: junta a rede neural, a memória de aprendizado, o MRP e a
alocação de capacidade em um ciclo de planejamento completo.

Ciclo:
  1. Retreina a rede neural com a demanda realizada até agora
  2. Prevê a demanda do horizonte
  3. Monta os parâmetros de cada item com o que já foi aprendido
  4. Roda o MRP (o que, quanto e quando produzir e comprar)
  5. Aloca as ordens de produção nos centros de trabalho (onde produzir)
Depois que a semana acontece, `registrar_*` alimenta o aprendizado.
"""

from dataclasses import dataclass

import numpy as np

from .aprendizado import MemoriaAprendizado
from .capacidade import alocar_capacidade
from .mrp import calcular_mrp
from .previsao import PrevisorDemanda, previsao_media_movel

PREMISSAS_PADRAO = {
    "horizonte_semanas": 12,
    "semanas_congeladas": 1,
    "nivel_servico": 0.95,
    "percentil_lead_time": 90,
    "usar_centro_alternativo": True,
    "amostras_minimas_aprendizado": 3,
    "suavizacao_aprendizado": 0.3,
    "epocas_treino_inicial": 300,
    "epocas_treino_por_ciclo": 20,
    "camadas_rede": [16],
}


@dataclass
class Plano:
    semana: int
    previsao: dict
    demanda_planejada: dict
    parametros: dict
    ordens: list
    tabela_mrp: dict
    carga_centros: dict
    capacidade_efetiva: dict
    usou_rede_neural: bool


class PlanejadorPCP:
    def __init__(self, itens, estrutura, capacidade_horas, premissas=None,
                 aprender=True, seed=0):
        self.itens = itens
        self.estrutura = estrutura
        self.capacidade_horas = capacidade_horas
        self.premissas = {**PREMISSAS_PADRAO, **(premissas or {})}
        self.aprender = aprender

        pais = {pai for pai, _, _ in estrutura}
        filhos = {filho for _, filho, _ in estrutura}
        self.acabados = [c for c in itens if c in pais and c not in filhos]

        self.previsor = PrevisorDemanda(self.acabados, self.premissas["camadas_rede"], seed)
        self.memoria = MemoriaAprendizado(
            amostras_minimas=self.premissas["amostras_minimas_aprendizado"],
            suavizacao=self.premissas["suavizacao_aprendizado"],
        )

    # ------------------------------------------------------------------
    def parametros(self):
        """Parâmetros de planejamento de cada item: aprendidos ou do cadastro."""
        p = self.premissas
        resultado = {}
        for codigo, item in self.itens.items():
            if not self.aprender:
                resultado[codigo] = {
                    "lead_time": item.lead_time,
                    "refugo": item.refugo,
                    "estoque_seguranca": item.estoque_seguranca,
                }
                continue
            lead_time = self.memoria.lead_time_planejamento(
                codigo, item.lead_time, p["percentil_lead_time"])
            seguranca = item.estoque_seguranca
            if codigo in self.acabados:
                seguranca = self.memoria.estoque_seguranca(
                    codigo, lead_time, p["nivel_servico"], item.estoque_seguranca)
            resultado[codigo] = {
                "lead_time": lead_time,
                "refugo": self.memoria.taxa_refugo(codigo, item.refugo),
                "estoque_seguranca": seguranca,
            }
        return resultado

    def eficiencias(self):
        if not self.aprender:
            return {c: 1.0 for c in self.capacidade_horas}
        return {c: self.memoria.eficiencia_centro(c) for c in self.capacidade_horas}

    # ------------------------------------------------------------------
    def planejar(self, semana, historico_demanda, estoque, recebimentos,
                 pedidos_firmes=None, pendencias=None):
        """
        historico_demanda: {produto acabado: demanda semanal até a semana anterior}
        estoque: {item: quantidade}
        recebimentos: {item: lista por semana do horizonte} (ordens em aberto)
        pedidos_firmes: {produto: lista por semana} (carteira de pedidos)
        pendencias: {produto: quantidade em atraso com o cliente}
        """
        p = self.premissas
        horizonte = p["horizonte_semanas"]

        usou_rede = False
        if self.aprender:
            self.previsor.treinar(
                historico_demanda, p["epocas_treino_inicial"], p["epocas_treino_por_ciclo"])
            previsao = self.previsor.prever(historico_demanda, horizonte)
            usou_rede = self.previsor.rede is not None
        else:
            previsao = {i: previsao_media_movel(historico_demanda[i], horizonte)
                        for i in self.acabados}

        # Demanda do plano = maior valor entre previsão e pedidos já firmes
        demanda = {}
        for produto in self.acabados:
            linha = np.array(previsao[produto], dtype=float)
            if pedidos_firmes and produto in pedidos_firmes:
                firmes = np.zeros(horizonte)
                valores = np.asarray(pedidos_firmes[produto], dtype=float)[:horizonte]
                firmes[:len(valores)] = valores
                linha = np.maximum(linha, firmes)
            if pendencias:
                linha[0] += pendencias.get(produto, 0.0)
            demanda[produto] = linha

        parametros = self.parametros()
        ordens, tabela = calcular_mrp(
            self.itens, self.estrutura, demanda, estoque, recebimentos,
            parametros, horizonte, p["semanas_congeladas"])

        eficiencia = self.eficiencias()
        carga = alocar_capacidade(
            ordens, self.itens, self.capacidade_horas, eficiencia, horizonte,
            p["usar_centro_alternativo"])
        capacidade_efetiva = {c: h * eficiencia[c] for c, h in self.capacidade_horas.items()}

        return Plano(semana, previsao, demanda, parametros, ordens, tabela,
                     carga, capacidade_efetiva, usou_rede)

    # ------------------------------------------------------------------
    # Retorno do chão de fábrica: é daqui que vem o aprendizado
    def registrar_demanda_real(self, plano, demanda_real):
        if not self.aprender:
            return
        for produto, real in demanda_real.items():
            self.memoria.registrar_previsao(produto, plano.previsao[produto][0], real)

    def registrar_recebimento_compra(self, item, semanas_lead_time):
        if self.aprender:
            self.memoria.registrar_lead_time(item, semanas_lead_time)

    def registrar_producao(self, item, quantidade, quantidade_boa):
        if self.aprender:
            self.memoria.registrar_producao(item, quantidade, quantidade_boa)

    def registrar_eficiencia(self, centro, eficiencia):
        if self.aprender:
            self.memoria.registrar_eficiencia(centro, eficiencia)
