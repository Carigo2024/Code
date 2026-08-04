"""
Motor de decisão de reabastecimento — 100% determinístico e auditável.

Para cada item:
  1. Estoque-base = disponível físico (Σ Qtde-Empenho, armazéns 01/05/22).
  2. Reprojeta o saldo MÊS A MÊS a partir desse estoque-base, usando as
     entradas programadas e as saídas do relatório (nunca médias).
  3. Ruptura = primeiro mês em que o saldo projetado cai abaixo do
     Estoque Segurança (piso vindo do próprio relatório, não recalculado).
  4. Volume = déficit para trazer o saldo mínimo de volta ao Estoque
     Segurança, respeitando Lote Mínimo e Lote Econômico.
  5. Lead time entra SÓ como urgência (não desloca a curva): urgente se a
     ruptura ocorre antes do que o lead time do item conseguiria repor.
  6. Cruza com a Necessidade nativa do Protheus e destaca divergências.
  7. Gera justificativa textual com os valores intermediários.

Todos os números que sustentam a decisão ficam expostos no objeto Decisao.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import datetime

from . import config
from .leitor_estoque import EstoqueSKU
from .leitor_mrp import ItemMRP


@dataclass
class Decisao:
    codigo: str
    descricao: str
    tipo: str
    unidade: str                  # unidade de medida (UN, G, ...)
    natureza: str                 # Comprar / Produzir / Verificar
    acao: str                     # texto do status
    volume: float                 # quanto pedir (0 se nada)
    urgente: bool
    revisao_manual: bool

    # valores intermediários (auditoria)
    estoque_base: float           # disponível físico usado na decisão
    estoque_mrp_m1: float         # +Estoque[m1] do relatório (conferência)
    divergencia_estoque: float    # base - mrp_m1
    estoque_seguranca: float
    lead_time: float
    lote_minimo: float
    lote_economico: float
    saldo_projetado: list[float] = field(default_factory=list)
    mes_ruptura_idx: int | None = None
    data_ruptura: datetime | None = None
    saldo_min: float = 0.0
    consumo_diario: float = 0.0
    cobertura_dias: float = 0.0
    # necessidade mensal time-phased (planned receipts por mês de disponibilidade)
    necessidade_mensal: list[float] = field(default_factory=list)

    # cruzamento com Protheus
    necessidade_protheus_total: float = 0.0
    primeiro_mes_nec_protheus: datetime | None = None
    divergente_do_protheus: bool = False

    alertas: list[str] = field(default_factory=list)
    justificativa: str = ""


def _aplicar_lote(q: float, lote_min: float, lote_econ: float) -> float:
    """Arredonda a quantidade para respeitar Lote Econômico (múltiplo) e Lote Mínimo (piso)."""
    if q <= 0:
        return 0.0
    if lote_econ > 0:
        q = math.ceil(q / lote_econ) * lote_econ
    if lote_min > 0 and q < lote_min:
        q = lote_min
    return q


def _fmt(x: float) -> str:
    return f"{x:,.0f}".replace(",", ".")


def decidir(
    item: ItemMRP,
    estoque: EstoqueSKU | None,
    periodos: list[datetime],
) -> Decisao:
    n = len(periodos)
    estoque_base = estoque.disponivel if estoque else 0.0
    mrp_m1 = item.serie(config.ROTULO_ESTOQUE, n)[0] if item.series else 0.0

    d = Decisao(
        codigo=item.codigo,
        descricao=item.descricao,
        tipo=item.tipo,
        unidade=item.unidade or "UN",
        natureza=config.TIPO_NATUREZA.get(item.tipo, "Verificar"),
        acao="",
        volume=0.0,
        urgente=False,
        revisao_manual=item.codigo in config.SKUS_REVISAO_MANUAL,
        estoque_base=estoque_base,
        estoque_mrp_m1=mrp_m1,
        divergencia_estoque=estoque_base - mrp_m1,
        estoque_seguranca=item.estoque_seguranca,
        lead_time=item.lead_time,
        lote_minimo=item.lote_minimo,
        lote_economico=item.lote_economico,
    )

    # --- séries mensais ---
    entradas = item.serie(config.ROTULO_ENTRADAS, n)
    saidas = [
        sum(sinal * item.serie(rot, n)[m] for rot, sinal in config.COMPONENTES_SAIDA.items())
        for m in range(n)
    ]
    nec_protheus = item.serie(config.ROTULO_NECESSIDADE, n)

    # --- reprojeção do saldo a partir do estoque físico ---
    saldo: list[float] = []
    corrente = estoque_base
    for m in range(n):
        corrente = corrente + entradas[m] - saidas[m]
        saldo.append(corrente)
    d.saldo_projetado = saldo
    d.saldo_min = min(saldo) if saldo else estoque_base

    # necessidade mensal time-phased: quanto precisa ficar DISPONÍVEL em cada mês
    # para manter o saldo >= estoque de segurança (planned receipts por período).
    receipts = [0.0] * n
    proj = estoque_base
    for m in range(n):
        proj = proj + entradas[m] - saidas[m]
        if proj < item.estoque_seguranca:
            need = _aplicar_lote(item.estoque_seguranca - proj, item.lote_minimo, item.lote_economico)
            receipts[m] = need
            proj += need
    d.necessidade_mensal = receipts

    # --- ruptura: 1º mês com saldo < estoque de segurança ---
    for m in range(n):
        if saldo[m] < item.estoque_seguranca:
            d.mes_ruptura_idx = m
            d.data_ruptura = periodos[m]
            break

    # --- consumo diário (run-rate de curto prazo) e cobertura em dias ---
    jan = min(config.MESES_RUN_RATE, n)
    consumo_curto = sum(max(saidas[m], 0.0) for m in range(jan))
    d.consumo_diario = consumo_curto / (jan * config.DIAS_POR_MES) if consumo_curto > 0 else 0.0
    if d.consumo_diario > 0:
        d.cobertura_dias = (estoque_base - item.estoque_seguranca) / d.consumo_diario
    else:
        d.cobertura_dias = float("inf")

    # --- cruzamento com a Necessidade nativa do Protheus ---
    d.necessidade_protheus_total = sum(nec_protheus)
    for m in range(n):
        if nec_protheus[m] > 0:
            d.primeiro_mes_nec_protheus = periodos[m]
            break

    # --- alerta de reconciliação de estoque ---
    if abs(d.divergencia_estoque) > config.LIMIAR_DIVERGENCIA_ESTOQUE:
        d.alertas.append(
            f"Estoque físico (01/05/22)={_fmt(estoque_base)} diverge do +Estoque[m1] "
            f"do MRP={_fmt(mrp_m1)} (dif {_fmt(d.divergencia_estoque)})."
        )
    if estoque is None:
        d.alertas.append("Sem estoque físico nos armazéns-alvo (base=0).")
    if estoque and estoque.qtd_apos_pull_date > 0:
        d.alertas.append(
            f"{_fmt(estoque.qtd_apos_pull_date)} un. já além do Pull Date — não consumir sem revisão."
        )

    # --- decisão de volume ---
    if d.natureza == "Verificar":
        d.acao = "Verificar cadastro (Tipo ausente/desconhecido)"
    elif d.mes_ruptura_idx is None:
        d.acao = "OK — coberto no horizonte"
        d.volume = 0.0
    else:
        deficit = item.estoque_seguranca - d.saldo_min
        d.volume = _aplicar_lote(deficit, item.lote_minimo, item.lote_economico)
        # urgência: a ruptura ocorre antes do que o lead time repõe?
        d.urgente = (d.mes_ruptura_idx * config.DIAS_POR_MES) <= item.lead_time
        verbo = d.natureza
        d.acao = f"{verbo}{' URGENTE' if d.urgente else ''}"

    # --- cruzamento com Protheus (vale para comprar E para não-comprar) ---
    if d.natureza != "Verificar":
        eu_ajo = d.volume > 0
        protheus_aponta = d.necessidade_protheus_total > 0
        if eu_ajo and protheus_aponta:
            base_cmp = max(d.necessidade_protheus_total, 1.0)
            if abs(d.volume - d.necessidade_protheus_total) / base_cmp > config.LIMIAR_DIVERGENCIA_NECESSIDADE:
                d.divergente_do_protheus = True
                d.alertas.append(
                    f"Volume calculado ({_fmt(d.volume)}) diverge >30% da Necessidade "
                    f"Protheus somada ({_fmt(d.necessidade_protheus_total)}) — revisar."
                )
        elif eu_ajo != protheus_aponta:
            d.divergente_do_protheus = True
            if protheus_aponta:
                d.alertas.append(
                    f"Eu concluo SEM pedido, mas o Protheus aponta Necessidade de "
                    f"{_fmt(d.necessidade_protheus_total)} un. — revisar (provável diferença de "
                    "estoque-base ou de ponto de pedido)."
                )
            else:
                d.alertas.append(
                    f"Eu recomendo pedido de {_fmt(d.volume)} un., mas o Protheus não aponta "
                    "Necessidade — revisar."
                )

    # --- exceção: revisão manual (importados de contrato) ---
    if d.revisao_manual:
        d.acao = f"REVISÃO MANUAL — {config.SKUS_REVISAO_MANUAL[item.codigo]}"
        d.alertas.insert(0, "Item de contrato anual/embarque bimestral: NÃO decidir automaticamente.")

    d.justificativa = _montar_justificativa(d, periodos)
    return d


def _montar_justificativa(d: Decisao, periodos: list[datetime]) -> str:
    """Justificativa determinística por template, usando os valores intermediários."""
    cob = "∞" if d.cobertura_dias == float("inf") else f"{d.cobertura_dias:.0f}d"
    partes: list[str] = []

    partes.append(
        f"Estoque disponível hoje (arm. 01/05/22, líquido de empenho) = {_fmt(d.estoque_base)} un.; "
        f"estoque de segurança = {_fmt(d.estoque_seguranca)} un."
    )

    if d.mes_ruptura_idx is None:
        partes.append(
            f"Projetando as entradas e saídas mês a mês, o saldo nunca cai abaixo do "
            f"estoque de segurança no horizonte (saldo mínimo = {_fmt(d.saldo_min)} un.). "
            f"Cobertura ≈ {cob}. Portanto, não é necessário pedido agora."
        )
    else:
        dr = d.data_ruptura.strftime("%m/%Y")
        partes.append(
            f"O saldo projetado cai abaixo do estoque de segurança pela primeira vez em "
            f"{dr} (saldo mínimo no horizonte = {_fmt(d.saldo_min)} un.). "
            f"Lead time do item = {d.lead_time:.0f} dias → "
            + ("um pedido feito hoje NÃO chega a tempo (urgente)."
               if d.urgente else "há prazo para repor dentro do lead time.")
        )
        partes.append(
            f"Volume = déficit para recompor o estoque de segurança "
            f"({_fmt(d.estoque_seguranca)} − ({_fmt(d.saldo_min)}) = {_fmt(d.estoque_seguranca - d.saldo_min)} un.), "
            f"ajustado a Lote Mín={_fmt(d.lote_minimo)} / Lote Econ={_fmt(d.lote_economico)} "
            f"→ {d.natureza.lower()} {_fmt(d.volume)} un."
        )

    if d.necessidade_protheus_total > 0 or d.divergente_do_protheus:
        mp = d.primeiro_mes_nec_protheus.strftime("%m/%Y") if d.primeiro_mes_nec_protheus else "-"
        conf = "DIVERGE da minha conclusão (revisar)" if d.divergente_do_protheus else "consistente com a minha conclusão"
        partes.append(
            f"Necessidade nativa do Protheus: {_fmt(d.necessidade_protheus_total)} un. "
            f"(1ª ocorrência em {mp}) — {conf}."
        )

    if d.revisao_manual:
        partes.append(
            "⚠️ Item importado de contrato anual com embarque bimestral e validade parcial: "
            "a recomendação acima é apenas indicativa — decidir manualmente."
        )
    return " ".join(partes)


def decidir_todos(
    itens: list[ItemMRP],
    estoques: dict[str, EstoqueSKU],
    periodos: list[datetime],
) -> list[Decisao]:
    return [decidir(it, estoques.get(it.codigo), periodos) for it in itens]
