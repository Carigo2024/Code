#!/usr/bin/env python3
"""
Refresh mensal da análise de reabastecimento (um comando).

Uso:
    python3 executar.py <relatorio_mrp.xlsx> <estoque.xlsx> [saida.xlsx]

Lê o relatório do MRP e o arquivo de estoque do mês, recalcula todas as
decisões e exporta o resultado. Se a estrutura de qualquer arquivo não bater
com o esperado, aborta com mensagem clara em vez de gerar dados errados.

Nomes de arquivo variam a cada mês — por isso são passados como argumento.
Uma GUI de "selecionar arquivos + botão atualizar" pode envolver esta função.
"""
from __future__ import annotations

import sys
from datetime import datetime

from mrp import (
    ErroEstruturaEstoque,
    ErroEstruturaMRP,
    decidir_todos,
    ler_estoque,
    ler_relatorio_mrp,
)
from mrp.exportador import exportar


def executar(caminho_mrp: str, caminho_estoque: str, caminho_saida: str) -> int:
    try:
        itens, periodos, avisos_mrp = ler_relatorio_mrp(caminho_mrp)
    except ErroEstruturaMRP as e:
        print(f"ERRO ao ler o relatório do MRP: {e}", file=sys.stderr)
        return 2
    try:
        estoques, avisos_est = ler_estoque(caminho_estoque)
    except ErroEstruturaEstoque as e:
        print(f"ERRO ao ler o arquivo de estoque: {e}", file=sys.stderr)
        return 2

    for a in avisos_mrp + avisos_est:
        print(f"AVISO: {a}")

    decisoes = decidir_todos(itens, estoques, periodos)
    exportar(decisoes, caminho_saida, periodos)

    urg = sum(d.urgente for d in decisoes)
    div = sum(d.divergente_do_protheus for d in decisoes)
    rev = sum(d.revisao_manual for d in decisoes)
    print(f"OK: {len(decisoes)} itens analisados | {urg} urgentes | "
          f"{div} divergentes do Protheus | {rev} revisão manual.")
    print(f"Resultado exportado em: {caminho_saida}")
    return 0


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print(__doc__)
        return 1
    caminho_mrp, caminho_estoque = argv[1], argv[2]
    saida = argv[3] if len(argv) > 3 else f"analise_reabastecimento_{datetime.now():%Y%m%d}.xlsx"
    return executar(caminho_mrp, caminho_estoque, saida)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
