"""
Previsão de demanda com a rede neural.

A rede aprende padrões da demanda de cada produto: tendência, sazonalidade
anual e o comportamento das últimas semanas. A cada ciclo ela é retreinada com
a demanda que realmente aconteceu e fica um pouco melhor.

Para funcionar com qualquer volume, cada série é dividida pela sua média
recente (escala). A rede aprende o "formato" da demanda, e o resultado é
multiplicado de volta pela escala.
"""

import numpy as np

from .rede_neural import RedeNeural

SEMANAS_ANO = 52
LAGS = (1, 2, 3, 4, SEMANAS_ANO)


class PrevisorDemanda:
    def __init__(self, itens, camadas_ocultas=(16,), seed=0):
        self.itens = list(itens)
        self.camadas_ocultas = camadas_ocultas
        self.seed = seed
        self.rede = None
        self.ciclos_treinados = 0

    @property
    def n_entradas(self):
        # lags + média móvel + sazonalidade (seno, cosseno) + identificação do item
        return len(LAGS) + 1 + 2 + len(self.itens)

    @staticmethod
    def historico_minimo():
        return max(LAGS) + 8

    # ------------------------------------------------------------------
    def _escala(self, serie, t):
        janela = serie[max(0, t - SEMANAS_ANO):t]
        return max(float(np.mean(janela)), 1e-6)

    def _entrada(self, item, serie, t):
        escala = self._escala(serie, t)
        lags = [serie[t - k] / escala for k in LAGS]
        media_movel = np.mean(serie[t - 4:t]) / escala
        angulo = 2 * np.pi * (t % SEMANAS_ANO) / SEMANAS_ANO
        id_item = [1.0 if i == item else 0.0 for i in self.itens]
        return np.array([*lags, media_movel, np.sin(angulo), np.cos(angulo), *id_item]), escala

    def _montar_treino(self, historico):
        X, y = [], []
        for item in self.itens:
            serie = np.asarray(historico[item], dtype=float)
            for t in range(max(LAGS), len(serie)):
                entrada, escala = self._entrada(item, serie, t)
                X.append(entrada)
                y.append(serie[t] / escala)
        return np.array(X), np.array(y)

    # ------------------------------------------------------------------
    def treinar(self, historico, epocas_iniciais=300, epocas_incrementais=20):
        """Treina a rede com o histórico. Nos ciclos seguintes, só ajusta."""
        if min(len(historico[i]) for i in self.itens) < self.historico_minimo():
            return None
        X, y = self._montar_treino(historico)
        if self.rede is None:
            self.rede = RedeNeural(self.n_entradas, self.camadas_ocultas, seed=self.seed)
            epocas = epocas_iniciais
        else:
            epocas = epocas_incrementais
        erros = self.rede.treinar(X, y, epocas=epocas)
        self.ciclos_treinados += 1
        return erros[-1]

    def prever(self, historico, horizonte):
        """Devolve {item: array com a previsão das próximas `horizonte` semanas}."""
        previsoes = {}
        for item in self.itens:
            serie = list(np.asarray(historico[item], dtype=float))
            if self.rede is None or len(serie) < self.historico_minimo():
                previsoes[item] = previsao_media_movel(serie, horizonte)
                continue
            futuro = []
            for _ in range(horizonte):
                t = len(serie)
                entrada, escala = self._entrada(item, np.array(serie), t)
                valor = max(0.0, float(self.rede.prever(entrada)[0]) * escala)
                futuro.append(valor)
                serie.append(valor)  # previsão recursiva: usa o que acabou de prever
            previsoes[item] = np.array(futuro)
        return previsoes


def previsao_media_movel(serie, horizonte, janela=4):
    """Método tradicional, usado como referência e quando falta histórico."""
    base = np.mean(serie[-janela:]) if len(serie) else 0.0
    return np.full(horizonte, float(base))
