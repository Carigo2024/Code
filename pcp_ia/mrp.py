"""
Motor de cálculo MRP (determinístico).

Aqui não tem "IA": são as contas clássicas do planejamento de materiais,
feitas com exatidão. A inteligência entra nos parâmetros que este motor
recebe (previsão, lead times, refugo, estoque de segurança), que vêm da rede
neural e da memória de aprendizado.

Para cada item, do produto acabado até a matéria-prima:
  necessidade bruta   = demanda independente + consumo pelos itens pai
  estoque projetado   = estoque + recebimentos - necessidade
  necessidade líquida = o que falta para manter o estoque de segurança
  ordem planejada     = necessidade líquida corrigida por refugo e regras de lote
  liberação da ordem  = data de necessidade - lead time
"""

import math
from dataclasses import dataclass, field


@dataclass
class Item:
    codigo: str
    descricao: str
    tipo: str                       # "fabricado" ou "comprado"
    lead_time: int                  # semanas (valor do cadastro)
    lote_minimo: float = 0
    lote_multiplo: float = 1
    estoque_seguranca: float = 0    # valor do cadastro
    refugo: float = 0.0             # valor do cadastro
    centro: str = ""                # centro de trabalho preferencial
    centro_alternativo: str = ""
    horas_por_unidade: float = 0.0
    fornecedor: str = ""
    custo: float = 0.0              # custo unitário, para valorizar o estoque


@dataclass
class OrdemPlanejada:
    item: str
    tipo: str                       # "produzir" ou "comprar"
    quantidade: float
    semana_liberacao: int
    semana_necessidade: int
    alertas: list = field(default_factory=list)
    centro: str = ""
    atraso: int = 0                 # semanas de atraso na liberação


def aplicar_lote(necessidade, item: Item):
    if necessidade <= 0:
        return 0.0
    quantidade = max(necessidade, item.lote_minimo)
    multiplo = item.lote_multiplo or 1
    return math.ceil(quantidade / multiplo - 1e-9) * multiplo


def niveis_estrutura(itens, estrutura):
    """Nível mais baixo de cada item na estrutura (low-level code)."""
    filhos = {}
    for pai, filho, _ in estrutura:
        filhos.setdefault(pai, []).append(filho)
    nivel = {codigo: 0 for codigo in itens}

    def descer(codigo, n, caminho):
        if codigo in caminho:
            raise ValueError(f"Estrutura circular envolvendo {codigo}")
        nivel[codigo] = max(nivel[codigo], n)
        for filho in filhos.get(codigo, []):
            descer(filho, n + 1, caminho | {codigo})

    for codigo in itens:
        descer(codigo, 0, frozenset())
    return nivel


def calcular_mrp(itens, estrutura, demanda_independente, estoque, recebimentos,
                 parametros, horizonte, semanas_congeladas=0):
    """
    itens: {codigo: Item}
    estrutura: [(pai, filho, quantidade_por_unidade)]
    demanda_independente: {codigo: lista por semana}
    estoque: {codigo: quantidade}
    recebimentos: {codigo: lista por semana} (ordens já abertas)
    parametros: {codigo: {"lead_time", "refugo", "estoque_seguranca"}} já aprendidos
    """
    nivel = niveis_estrutura(itens, estrutura)
    necessidade_bruta = {c: [0.0] * horizonte for c in itens}
    for codigo, demanda in demanda_independente.items():
        for t in range(min(horizonte, len(demanda))):
            necessidade_bruta[codigo][t] += float(demanda[t])

    ordens = []
    tabela = {}
    for codigo in sorted(itens, key=lambda c: nivel[c]):
        item = itens[codigo]
        p = parametros[codigo]
        lead_time = p["lead_time"]
        refugo = min(p["refugo"], 0.9)
        seguranca = p["estoque_seguranca"]
        receb = list(recebimentos.get(codigo, [0.0] * horizonte)) + [0.0] * horizonte

        disponivel = float(estoque.get(codigo, 0.0))
        linhas = []
        for t in range(horizonte):
            disponivel += receb[t] - necessidade_bruta[codigo][t]
            liquida = max(0.0, seguranca - disponivel)
            recebimento_planejado = 0.0
            if liquida > 0:
                # Pede a mais para compensar o refugo esperado
                quantidade = aplicar_lote(liquida / (1 - refugo), item)
                recebimento_planejado = quantidade * (1 - refugo)
                disponivel += recebimento_planejado
                liberacao = t - lead_time
                ordem = OrdemPlanejada(
                    item=codigo,
                    tipo="produzir" if item.tipo == "fabricado" else "comprar",
                    quantidade=quantidade,
                    semana_liberacao=max(0, liberacao),
                    semana_necessidade=t,
                    centro=item.centro,
                )
                atrasada = next((o for o in ordens if o.item == codigo and o.atraso > 0), None)
                if liberacao < 0 and atrasada:
                    # Junta todas as necessidades atrasadas do item numa única ordem urgente
                    atrasada.quantidade += quantidade
                    atrasada.atraso = max(atrasada.atraso, -liberacao)
                    atrasada.alertas = [
                        f"ATRASADA: deveria ter sido liberada há {atrasada.atraso} semana(s)"]
                    ordem = atrasada
                else:
                    if liberacao < 0:
                        ordem.atraso = -liberacao
                        ordem.alertas.append(
                            f"ATRASADA: deveria ter sido liberada há {-liberacao} semana(s)")
                    elif t < semanas_congeladas:
                        ordem.alertas.append("Necessidade dentro do horizonte congelado: exige aprovação")
                    ordens.append(ordem)
                # A liberação da ordem gera necessidade nos componentes
                for pai, filho, qtd in estrutura:
                    if pai == codigo:
                        necessidade_bruta[filho][max(0, liberacao)] += quantidade * qtd
            linhas.append({
                "semana": t,
                "necessidade_bruta": necessidade_bruta[codigo][t],
                "recebimentos": receb[t],
                "estoque_projetado": disponivel,
                "recebimento_planejado": recebimento_planejado,
            })
        tabela[codigo] = linhas
    return ordens, tabela
