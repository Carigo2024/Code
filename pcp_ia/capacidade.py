"""
Onde produzir: aloca as ordens de produção nos centros de trabalho.

A capacidade de cada centro é multiplicada pela eficiência APRENDIDA (e não
pela do cadastro), então o plano passa a respeitar o que a fábrica realmente
consegue produzir. Quando o centro preferencial lota, o que não cabe vai para
o centro alternativo (a ordem é dividida). Se ainda sobrar, a ordem recebe um
alerta de sobrecarga.
"""

import math
from dataclasses import replace


def alocar_capacidade(ordens, itens, capacidade_horas, eficiencia, horizonte,
                      usar_alternativo=True):
    """
    ordens: lista de OrdemPlanejada; as de produção são alocadas e, se
            necessário, divididas entre centros (a lista é alterada no lugar)
    capacidade_horas: {centro: horas disponíveis por semana}
    eficiencia: {centro: eficiência aprendida (0-1)}
    Devolve {centro: [horas carregadas por semana]}.
    """
    livre = {c: [h * eficiencia.get(c, 1.0)] * horizonte for c, h in capacidade_horas.items()}
    carga = {c: [0.0] * horizonte for c in capacidade_horas}

    def carregar(ordem, centro, t):
        horas = ordem.quantidade * itens[ordem.item].horas_por_unidade
        ordem.centro = centro
        livre[centro][t] -= horas
        carga[centro][t] += horas

    producao = [o for o in ordens if o.tipo == "produzir"]
    # Prioridade: quem é necessário antes é alocado primeiro
    for ordem in sorted(producao, key=lambda o: (o.semana_necessidade, o.semana_liberacao)):
        item = itens[ordem.item]
        t = min(ordem.semana_liberacao, horizonte - 1)
        preferencial = item.centro
        alternativo = item.centro_alternativo if usar_alternativo else ""

        cabe = _quantidade_que_cabe(livre[preferencial][t], item)
        if cabe >= ordem.quantidade or not alternativo:
            carregar(ordem, preferencial, t)
        elif cabe > 0:
            # Divide: parte no preferencial, o restante no alternativo
            resto = replace(ordem, quantidade=ordem.quantidade - cabe, alertas=list(ordem.alertas))
            ordem.quantidade = cabe
            carregar(ordem, preferencial, t)
            resto.alertas.append(f"Parte da ordem desviada para {alternativo}")
            ordens.append(resto)
            carregar(resto, alternativo, t)
            ordem = resto
        else:
            ordem.alertas.append(f"Desviada para centro alternativo {alternativo}")
            carregar(ordem, alternativo, t)

        if livre[ordem.centro][t] < -1e-6:
            ordem.alertas.append(
                f"SOBRECARGA em {ordem.centro} na semana {t}: faltam "
                f"{-livre[ordem.centro][t]:.1f} h (antecipar, hora extra ou renegociar prazo)")
    return carga


def _quantidade_que_cabe(horas_livres, item):
    if item.horas_por_unidade <= 0:
        return math.inf
    if horas_livres <= 0:
        return 0
    multiplo = item.lote_multiplo or 1
    return math.floor(horas_livres / item.horas_por_unidade / multiplo) * multiplo
