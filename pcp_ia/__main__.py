"""
Linha de comando da IA de PCP.

  python -m pcp_ia simular  [--ciclos 52]          compara planejar COM e SEM aprendizado
  python -m pcp_ia exemplo  [--pasta dados_exemplo] gera planilhas de exemplo para preencher
  python -m pcp_ia planejar --dados PASTA [--saida PASTA]  roda um ciclo com os seus dados
"""

import argparse
import os

import numpy as np

from .dados import escrever_csv, gerar_pasta_exemplo, planejar_pasta
from .relatorio import salvar_plano
from .simulacao import simular


def comando_simular(args):
    print(f"Simulando {args.ciclos} semanas na empresa fictícia...\n")
    sem = simular(ciclos=args.ciclos, aprender=False, seed=args.seed)
    com = simular(ciclos=args.ciclos, aprender=True, seed=args.seed)

    bloco = max(1, args.ciclos // 4)
    print(f"{'Semanas':>10} | {'Erro previsão':^19} | {'Nível de serviço':^19} | "
          f"{'Pendências (un.)':^19} | {'Estoque (R$ mil)':^19}")
    print(f"{'':>10} | {'sem IA':>9} {'com IA':>9} | {'sem IA':>9} {'com IA':>9} | "
          f"{'sem IA':>9} {'com IA':>9} | {'sem IA':>9} {'com IA':>9}")
    print("-" * 98)
    for inicio in range(0, args.ciclos, bloco):
        def media(res, chave):
            return np.mean([m[chave] for m in res.metricas[inicio:inicio + bloco]])
        fim = min(inicio + bloco, args.ciclos)
        print(f"{inicio + 1:>4} a {fim:<3} | "
              f"{media(sem, 'erro_previsao_pct'):>8.1f}% {media(com, 'erro_previsao_pct'):>8.1f}% | "
              f"{media(sem, 'nivel_servico_pct'):>8.1f}% {media(com, 'nivel_servico_pct'):>8.1f}% | "
              f"{media(sem, 'pendencias_unid'):>9.0f} {media(com, 'pendencias_unid'):>9.0f} | "
              f"{media(sem, 'valor_estoque') / 1000:>9.1f} {media(com, 'valor_estoque') / 1000:>9.1f}")

    os.makedirs(args.saida, exist_ok=True)
    escrever_csv(os.path.join(args.saida, "metricas_sem_aprendizado.csv"), sem.metricas)
    escrever_csv(os.path.join(args.saida, "metricas_com_aprendizado.csv"), com.metricas)
    salvar_plano(com.ultimo_plano, com.planejador, args.saida)
    print(f"\nMétricas e plano da última semana salvos em: {args.saida}/")


def comando_exemplo(args):
    gerar_pasta_exemplo(args.pasta)
    print(f"Pasta de exemplo criada em: {args.pasta}/")
    print(f"Agora rode:  python -m pcp_ia planejar --dados {args.pasta}")


def comando_planejar(args):
    planejador, plano = planejar_pasta(args.dados)
    saida = args.saida or os.path.join(args.dados, "saida")
    texto = salvar_plano(plano, planejador, saida)
    print(texto)
    print(f"\nArquivos gerados em: {saida}/ (ordens_planejadas.csv, previsao_demanda.csv, "
          "carga_centros.csv, relatorio.md)")


def main(argv=None):
    parser = argparse.ArgumentParser(prog="pcp_ia", description="IA de Planejamento e Controle da Produção")
    sub = parser.add_subparsers(dest="comando", required=True)

    p = sub.add_parser("simular", help="demonstra o aprendizado na empresa fictícia")
    p.add_argument("--ciclos", type=int, default=52)
    p.add_argument("--seed", type=int, default=42)
    p.add_argument("--saida", default="saida_simulacao")
    p.set_defaults(func=comando_simular)

    p = sub.add_parser("exemplo", help="gera planilhas de exemplo no formato esperado")
    p.add_argument("--pasta", default="dados_exemplo")
    p.set_defaults(func=comando_exemplo)

    p = sub.add_parser("planejar", help="roda um ciclo de planejamento com os seus dados")
    p.add_argument("--dados", required=True)
    p.add_argument("--saida")
    p.set_defaults(func=comando_planejar)

    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
