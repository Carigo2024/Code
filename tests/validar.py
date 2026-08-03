"""
Roda o pipeline completo nos arquivos de referência (dados_ref/) e imprime um
diagnóstico: estrutura lida, avisos, distribuição de decisões e uma amostra
detalhada — incluindo os 2 SKUs de revisão manual.

Uso: python3 -m tests.validar
"""
import sys
from collections import Counter

from mrp import ler_relatorio_mrp, ler_estoque, decidir_todos
from mrp import config

MRP = "dados_ref/mrp_ref.xlsx"
EST = "dados_ref/estoque_ref.xlsx"


def main() -> int:
    print("=== LEITURA MRP ===")
    itens, periodos, avisos_mrp = ler_relatorio_mrp(MRP)
    print(f"Itens lidos: {len(itens)} | meses: {len(periodos)} "
          f"({periodos[0]:%m/%Y}..{periodos[-1]:%m/%Y})")
    for a in avisos_mrp:
        print("  aviso:", a)

    print("\n=== LEITURA ESTOQUE (arm. 01/05/22, empenho consumido) ===")
    estoques, avisos_est = ler_estoque(EST)
    print(f"SKUs com estoque nos armazéns-alvo: {len(estoques)}")
    for a in avisos_est:
        print("  aviso:", a)

    print("\n=== DECISÕES ===")
    decisoes = decidir_todos(itens, estoques, periodos)
    print("Ações:", dict(Counter(d.acao.split(' —')[0].split(' (')[0] for d in decisoes)))
    print("Natureza:", dict(Counter(d.natureza for d in decisoes)))
    print("Urgentes:", sum(d.urgente for d in decisoes))
    print("Divergentes do Protheus:", sum(d.divergente_do_protheus for d in decisoes))
    print("Com alerta de reconciliação de estoque:",
          sum(any("diverge do +Estoque" in a for a in d.alertas) for d in decisoes))

    print("\n=== SKUs DE REVISÃO MANUAL ===")
    for d in decisoes:
        if d.revisao_manual:
            print(f"\n[{d.codigo}] {d.descricao} | Tipo={d.tipo} natureza={d.natureza}")
            print(f"  Ação: {d.acao}")
            print(f"  Justificativa: {d.justificativa}")

    print("\n=== AMOSTRA (3 comprar + 2 produzir + 2 coberto) ===")
    def amostra(filtro, k):
        return [d for d in decisoes if filtro(d)][:k]
    sel = (amostra(lambda d: d.natureza == "Comprar" and d.volume > 0 and not d.revisao_manual, 3)
           + amostra(lambda d: d.natureza == "Produzir" and d.volume > 0 and not d.revisao_manual, 2)
           + amostra(lambda d: d.mes_ruptura_idx is None, 2))
    for d in sel:
        print(f"\n[{d.codigo}] {d.descricao[:32]} | {d.acao} | vol={d.volume:.0f}")
        print(f"  {d.justificativa}")
        for a in d.alertas:
            print(f"  ! {a}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
