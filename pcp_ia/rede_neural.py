"""
Rede neural (perceptron multicamadas) implementada do zero com numpy.

Por que do zero? Para que você possa abrir o arquivo e ver exatamente o que a
rede faz: camadas densas, ativação ReLU, otimizador Adam e erro quadrático
médio. Não há nenhuma "caixa-preta" de biblioteca.

A rede suporta treino incremental (warm start): a cada ciclo de planejamento
ela continua de onde parou, ajustando os pesos com os dados novos em vez de
começar do zero.
"""

import numpy as np


class RedeNeural:
    def __init__(self, n_entradas, camadas_ocultas=(32, 16), taxa_aprendizado=1e-3, seed=0):
        self.n_entradas = n_entradas
        self.camadas_ocultas = tuple(camadas_ocultas)
        self.taxa = taxa_aprendizado
        self.rng = np.random.default_rng(seed)

        tamanhos = [n_entradas, *self.camadas_ocultas, 1]
        self.pesos = []
        self.vieses = []
        for entrada, saida in zip(tamanhos[:-1], tamanhos[1:]):
            # Inicialização He: adequada para ReLU
            self.pesos.append(self.rng.normal(0, np.sqrt(2 / entrada), size=(entrada, saida)))
            self.vieses.append(np.zeros(saida))

        # Estado do otimizador Adam (momentos de 1a e 2a ordem)
        self._m = [np.zeros_like(p) for p in self.pesos + self.vieses]
        self._v = [np.zeros_like(p) for p in self.pesos + self.vieses]
        self._passo = 0

        # Normalização das entradas e da saída (aprendida nos dados de treino)
        self.x_media = np.zeros(n_entradas)
        self.x_desvio = np.ones(n_entradas)
        self.y_media = 0.0
        self.y_desvio = 1.0
        self.treinada = False

    # ------------------------------------------------------------------
    def _propagar(self, X):
        """Passa X pela rede e devolve a saída e as ativações de cada camada."""
        ativacoes = [X]
        a = X
        for i, (W, b) in enumerate(zip(self.pesos, self.vieses)):
            z = a @ W + b
            a = z if i == len(self.pesos) - 1 else np.maximum(z, 0)  # ReLU nas ocultas
            ativacoes.append(a)
        return a[:, 0], ativacoes

    def prever(self, X):
        X = np.atleast_2d(np.asarray(X, dtype=float))
        Xn = (X - self.x_media) / self.x_desvio
        saida, _ = self._propagar(Xn)
        return saida * self.y_desvio + self.y_media

    # ------------------------------------------------------------------
    def treinar(self, X, y, epocas=200, tamanho_lote=32):
        """Treina (ou continua treinando) a rede. Devolve o erro por época."""
        X = np.atleast_2d(np.asarray(X, dtype=float))
        y = np.asarray(y, dtype=float)

        self.x_media = X.mean(axis=0)
        self.x_desvio = X.std(axis=0) + 1e-8
        self.y_media = float(y.mean())
        self.y_desvio = float(y.std() + 1e-8)
        Xn = (X - self.x_media) / self.x_desvio
        yn = (y - self.y_media) / self.y_desvio

        historico_erro = []
        n = len(Xn)
        for _ in range(epocas):
            ordem = self.rng.permutation(n)
            for inicio in range(0, n, tamanho_lote):
                idx = ordem[inicio:inicio + tamanho_lote]
                self._passo_gradiente(Xn[idx], yn[idx])
            previsto, _ = self._propagar(Xn)
            historico_erro.append(float(np.mean((previsto - yn) ** 2)))
        self.treinada = True
        return historico_erro

    def _passo_gradiente(self, X, y):
        saida, ativacoes = self._propagar(X)
        n = len(X)
        # Derivada do erro quadrático médio
        delta = ((saida - y) * 2 / n)[:, None]

        grad_W = [None] * len(self.pesos)
        grad_b = [None] * len(self.vieses)
        for i in reversed(range(len(self.pesos))):
            grad_W[i] = ativacoes[i].T @ delta
            grad_b[i] = delta.sum(axis=0)
            if i > 0:
                delta = (delta @ self.pesos[i].T) * (ativacoes[i] > 0)

        self._adam(grad_W + grad_b)

    def _adam(self, gradientes, beta1=0.9, beta2=0.999, eps=1e-8):
        self._passo += 1
        parametros = self.pesos + self.vieses
        for k, (p, g) in enumerate(zip(parametros, gradientes)):
            self._m[k] = beta1 * self._m[k] + (1 - beta1) * g
            self._v[k] = beta2 * self._v[k] + (1 - beta2) * g ** 2
            m_hat = self._m[k] / (1 - beta1 ** self._passo)
            v_hat = self._v[k] / (1 - beta2 ** self._passo)
            p -= self.taxa * m_hat / (np.sqrt(v_hat) + eps)

    # ------------------------------------------------------------------
    def salvar(self, caminho):
        dados = {
            "config": np.array([self.n_entradas, *self.camadas_ocultas]),
            "normalizacao": np.array([self.y_media, self.y_desvio]),
            "x_media": self.x_media,
            "x_desvio": self.x_desvio,
            "passo": np.array([self._passo]),
        }
        for k, p in enumerate(self.pesos + self.vieses):
            dados[f"param_{k}"] = p
        for k, (m, v) in enumerate(zip(self._m, self._v)):
            dados[f"m_{k}"] = m
            dados[f"v_{k}"] = v
        np.savez(caminho, **dados)

    @classmethod
    def carregar(cls, caminho, taxa_aprendizado=1e-3, seed=0):
        dados = np.load(caminho)
        config = dados["config"].tolist()
        rede = cls(config[0], config[1:], taxa_aprendizado, seed)
        n = len(rede.pesos)
        for k in range(2 * n):
            alvo = rede.pesos if k < n else rede.vieses
            alvo[k % n] = dados[f"param_{k}"]
            rede._m[k] = dados[f"m_{k}"]
            rede._v[k] = dados[f"v_{k}"]
        rede.y_media, rede.y_desvio = dados["normalizacao"].tolist()
        rede.x_media = dados["x_media"]
        rede.x_desvio = dados["x_desvio"]
        rede._passo = int(dados["passo"][0])
        rede.treinada = True
        return rede
