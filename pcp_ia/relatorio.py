"""
Saídas do plano: planilhas CSV (para importar no ERP) e um relatório em
Markdown explicando o que fazer e por quê.
"""

import os

from .dados import escrever_csv


def salvar_plano(plano, planejador, pasta_saida):
    os.makedirs(pasta_saida, exist_ok=True)
    itens = planejador.itens
    s0 = plano.semana

    escrever_csv(os.path.join(pasta_saida, "ordens_planejadas.csv"), [
        {
            "acao": o.tipo,
            "item": o.item,
            "descricao": itens[o.item].descricao,
            "quantidade": round(o.quantidade, 2),
            "semana_liberacao": s0 + o.semana_liberacao,
            "semana_necessidade": s0 + o.semana_necessidade,
            "onde": o.centro if o.tipo == "produzir" else itens[o.item].fornecedor,
            "alertas": " | ".join(o.alertas),
        }
        for o in sorted(plano.ordens, key=lambda o: (o.semana_liberacao, o.tipo, o.item))
    ], ["acao", "item", "descricao", "quantidade", "semana_liberacao",
        "semana_necessidade", "onde", "alertas"])

    escrever_csv(os.path.join(pasta_saida, "previsao_demanda.csv"), [
        {"item": p, "semana": s0 + t, "previsao": round(float(v), 1),
         "demanda_planejada": round(float(plano.demanda_planejada[p][t]), 1)}
        for p, serie in plano.previsao.items() for t, v in enumerate(serie)
    ])

    escrever_csv(os.path.join(pasta_saida, "carga_centros.csv"), [
        {"centro": c, "semana": s0 + t, "carga_horas": round(h, 1),
         "capacidade_efetiva_horas": round(plano.capacidade_efetiva[c], 1),
         "ocupacao_pct": round(100 * h / plano.capacidade_efetiva[c], 1)
         if plano.capacidade_efetiva[c] else 0}
        for c, horas in plano.carga_centros.items() for t, h in enumerate(horas)
    ])

    texto = gerar_relatorio(plano, planejador)
    with open(os.path.join(pasta_saida, "relatorio.md"), "w", encoding="utf-8") as f:
        f.write(texto)
    return texto


def _n(valor):
    """Número no formato brasileiro: 1.234"""
    return f"{valor:,.0f}".replace(",", ".")


def gerar_relatorio(plano, planejador):
    itens = planejador.itens
    s0 = plano.semana
    L = [f"# Plano de PCP da semana {s0}", ""]

    metodo = "rede neural" if plano.usou_rede_neural else "média móvel (histórico curto)"
    L += [f"Previsão de demanda feita por **{metodo}**, horizonte de "
          f"{len(next(iter(plano.previsao.values())))} semanas.", ""]

    agora = [o for o in plano.ordens if o.semana_liberacao == 0]
    L += ["## 1. Liberar ESTA semana", ""]
    for tipo, titulo in (("comprar", "Compras"), ("produzir", "Produção")):
        ordens = sorted([o for o in agora if o.tipo == tipo], key=lambda o: o.semana_necessidade)
        L += [f"### {titulo}", ""]
        if not ordens:
            L += ["Nada a liberar.", ""]
            continue
        onde = "Fornecedor" if tipo == "comprar" else "Centro"
        L += [f"| Item | Descrição | Quantidade | Precisa em | {onde} | Alertas |",
              "|---|---|---:|---|---|---|"]
        for o in ordens:
            local = itens[o.item].fornecedor if tipo == "comprar" else o.centro
            L.append(f"| {o.item} | {itens[o.item].descricao} | {_n(o.quantidade)} | "
                     f"sem. {s0 + o.semana_necessidade} | {local} | {'; '.join(o.alertas) or '-'} |")
        L.append("")

    alertas = [o for o in plano.ordens if any(
        a.startswith(("ATRASADA", "SOBRECARGA")) for a in o.alertas)]
    L += ["## 2. Pontos de atenção", ""]
    if alertas:
        for o in alertas[:20]:
            L.append(f"- **{o.item}** ({o.tipo}, {_n(o.quantidade)} un., precisa na semana "
                     f"{s0 + o.semana_necessidade}): {'; '.join(o.alertas)}")
    else:
        L.append("Nenhuma ordem atrasada ou centro sobrecarregado no horizonte.")
    L.append("")

    L += ["## 3. Previsão de demanda (próximas semanas)", "",
          "| Produto | " + " | ".join(f"sem. {s0 + t}" for t in range(6)) + " |",
          "|---|" + "---:|" * 6]
    for p, serie in plano.demanda_planejada.items():
        L.append(f"| {p} | " + " | ".join(f"{_n(v)}" for v in serie[:6]) + " |")
    L.append("")

    L += ["## 4. O que a IA aprendeu (cadastro × realidade)", "",
          "| Item | Lead time cadastro | Lead time usado | Refugo cadastro | Refugo aprendido | "
          "Est. segurança cadastro | Est. segurança usado |",
          "|---|---:|---:|---:|---:|---:|---:|"]
    for codigo, p in plano.parametros.items():
        item = itens[codigo]
        L.append(f"| {codigo} | {item.lead_time} | {p['lead_time']} | {item.refugo:.1%} | "
                 f"{p['refugo']:.1%} | {_n(item.estoque_seguranca)} | {_n(p['estoque_seguranca'])} |")
    L.append("")

    L += ["| Centro | Capacidade nominal (h/sem) | Eficiência aprendida | Capacidade efetiva | "
          "Ocupação esta semana |", "|---|---:|---:|---:|---:|"]
    for c, horas in planejador.capacidade_horas.items():
        efetiva = plano.capacidade_efetiva[c]
        ocupacao = plano.carga_centros[c][0] / efetiva if efetiva else 0
        L.append(f"| {c} | {_n(horas)} | {efetiva / horas if horas else 0:.0%} | "
                 f"{_n(efetiva)} | {ocupacao:.0%} |")
    L.append("")
    return "\n".join(L)
