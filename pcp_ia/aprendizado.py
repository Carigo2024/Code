"""
Memória de aprendizado: compara o que foi planejado com o que aconteceu.

A cada ciclo o sistema registra os resultados reais e corrige os parâmetros
que o cadastro do ERP costuma ter errados:

  * lead time real de cada item comprado (fornecedor atrasa?)
  * refugo real de cada item fabricado (quanto se perde na produção?)
  * eficiência real de cada centro de trabalho (a máquina rende 100%?)
  * erro da previsão de demanda (quanto estoque de segurança é preciso?)

Enquanto não há observações suficientes, vale o valor do cadastro.
"""

import json
import math

import numpy as np

# Fator z da distribuição normal para cada nível de serviço
TABELA_Z = {0.80: 0.84, 0.85: 1.04, 0.90: 1.28, 0.95: 1.65, 0.975: 1.96, 0.99: 2.33}


def fator_z(nivel_servico):
    niveis = sorted(TABELA_Z)
    mais_proximo = min(niveis, key=lambda n: abs(n - nivel_servico))
    return TABELA_Z[mais_proximo]


class MemoriaAprendizado:
    def __init__(self, amostras_minimas=3, suavizacao=0.3, janela_erros=26):
        self.amostras_minimas = amostras_minimas
        self.suavizacao = suavizacao
        self.janela_erros = janela_erros
        self.lead_times = {}   # item -> [lead times observados]
        self.refugo = {}       # item -> taxa suavizada
        self.eficiencia = {}   # centro -> eficiência suavizada
        self.erros = {}        # item -> [real - previsto]

    # ---------------------------------------------------------- registros
    def registrar_lead_time(self, item, semanas):
        self.lead_times.setdefault(item, []).append(int(semanas))

    def registrar_producao(self, item, quantidade_planejada, quantidade_boa):
        if quantidade_planejada <= 0:
            return
        taxa = 1 - quantidade_boa / quantidade_planejada
        self.refugo[item] = self._suavizar(self.refugo.get(item), taxa)

    def registrar_eficiencia(self, centro, eficiencia_observada):
        self.eficiencia[centro] = self._suavizar(self.eficiencia.get(centro), eficiencia_observada)

    def registrar_previsao(self, item, previsto, real):
        erros = self.erros.setdefault(item, [])
        erros.append(float(real) - float(previsto))
        del erros[:-self.janela_erros]

    def _suavizar(self, anterior, novo):
        if anterior is None:
            return float(novo)
        return (1 - self.suavizacao) * anterior + self.suavizacao * float(novo)

    # ---------------------------------------------------------- consultas
    def lead_time_planejamento(self, item, cadastro, percentil=90):
        """Lead time a usar no plano: um percentil alto do que foi observado."""
        observados = self.lead_times.get(item, [])
        if len(observados) < self.amostras_minimas:
            return cadastro
        return int(math.ceil(np.percentile(observados, percentil)))

    def taxa_refugo(self, item, cadastro=0.0):
        return self.refugo.get(item, cadastro)

    def eficiencia_centro(self, centro, cadastro=1.0):
        return self.eficiencia.get(centro, cadastro)

    def desvio_previsao(self, item):
        erros = self.erros.get(item, [])
        if len(erros) < self.amostras_minimas:
            return None
        return float(np.std(erros, ddof=1))

    def estoque_seguranca(self, item, lead_time, nivel_servico, cadastro=0.0):
        """Estoque de segurança = z * desvio do erro de previsão * raiz(lead time + 1)."""
        desvio = self.desvio_previsao(item)
        if desvio is None:
            return cadastro
        return fator_z(nivel_servico) * desvio * math.sqrt(lead_time + 1)

    # ---------------------------------------------------------- persistência
    def para_dict(self):
        return {
            "lead_times": self.lead_times,
            "refugo": self.refugo,
            "eficiencia": self.eficiencia,
            "erros": self.erros,
        }

    def salvar(self, caminho):
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(self.para_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def carregar(cls, caminho, **kwargs):
        memoria = cls(**kwargs)
        with open(caminho, encoding="utf-8") as f:
            dados = json.load(f)
        memoria.lead_times = dados.get("lead_times", {})
        memoria.refugo = dados.get("refugo", {})
        memoria.eficiencia = dados.get("eficiencia", {})
        memoria.erros = dados.get("erros", {})
        return memoria
