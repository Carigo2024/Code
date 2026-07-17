"""
Análise e visualização do comportamento de outliers.

Este script gera dados sintéticos com alguns valores atípicos (outliers),
identifica esses outliers usando dois métodos clássicos (IQR e Z-score) e
produz um gráfico com quatro painéis que ilustram o comportamento deles:

  1. Dispersão dos dados destacando os outliers
  2. Boxplot (evidencia visualmente os limites de outlier)
  3. Histograma com a curva de densidade
  4. Comparação dos limites IQR vs Z-score

Uso:
    python outliers_graph.py
    python outliers_graph.py --n 500 --output meu_grafico.png
"""

import argparse

import numpy as np
import matplotlib.pyplot as plt


def gerar_dados(n=200, seed=42):
    """Gera dados normais com alguns outliers injetados propositalmente."""
    rng = np.random.default_rng(seed)
    dados = rng.normal(loc=50, scale=8, size=n)

    # Injeta outliers para tornar o comportamento visível
    n_outliers = max(5, n // 40)
    posicoes = rng.choice(n, size=n_outliers, replace=False)
    dados[posicoes] = rng.uniform(low=90, high=120, size=n_outliers)
    # Alguns outliers baixos também
    posicoes_baixos = rng.choice(n, size=n_outliers // 2 + 1, replace=False)
    dados[posicoes_baixos] = rng.uniform(low=0, high=15, size=len(posicoes_baixos))

    return dados


def limites_iqr(dados, k=1.5):
    """Retorna os limites inferior e superior pelo método IQR (Tukey)."""
    q1, q3 = np.percentile(dados, [25, 75])
    iqr = q3 - q1
    return q1 - k * iqr, q3 + k * iqr


def limites_zscore(dados, k=3.0):
    """Retorna os limites inferior e superior pelo método Z-score."""
    media, desvio = np.mean(dados), np.std(dados)
    return media - k * desvio, media + k * desvio


def plotar(dados, output="outliers_comportamento.png"):
    inf_iqr, sup_iqr = limites_iqr(dados)
    inf_z, sup_z = limites_zscore(dados)

    mascara_outlier = (dados < inf_iqr) | (dados > sup_iqr)
    normais = dados[~mascara_outlier]
    outliers = dados[mascara_outlier]

    fig, axs = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle(
        "Comportamento dos Outliers", fontsize=18, fontweight="bold"
    )

    cor_normal = "#2c7fb8"
    cor_outlier = "#d7191c"

    # --- Painel 1: dispersão ---
    ax = axs[0, 0]
    indices = np.arange(len(dados))
    ax.scatter(
        indices[~mascara_outlier], normais,
        s=25, color=cor_normal, alpha=0.7, label="Normais",
    )
    ax.scatter(
        indices[mascara_outlier], outliers,
        s=60, color=cor_outlier, edgecolor="black",
        label=f"Outliers ({len(outliers)})", zorder=3,
    )
    ax.axhline(sup_iqr, color="gray", linestyle="--", linewidth=1)
    ax.axhline(inf_iqr, color="gray", linestyle="--", linewidth=1)
    ax.set_title("Dispersão dos dados (limites IQR tracejados)")
    ax.set_xlabel("Índice da amostra")
    ax.set_ylabel("Valor")
    ax.legend()

    # --- Painel 2: boxplot ---
    ax = axs[0, 1]
    bp = ax.boxplot(
        dados, orientation="vertical", patch_artist=True,
        flierprops=dict(marker="o", markerfacecolor=cor_outlier,
                        markersize=7, markeredgecolor="black"),
    )
    bp["boxes"][0].set_facecolor("#a6bddb")
    ax.set_title("Boxplot (pontos = outliers detectados)")
    ax.set_ylabel("Valor")
    ax.set_xticklabels(["Amostra"])

    # --- Painel 3: histograma + densidade ---
    ax = axs[1, 0]
    ax.hist(dados, bins=30, color=cor_normal, alpha=0.6,
            density=True, edgecolor="white")
    xs = np.linspace(dados.min(), dados.max(), 300)
    media, desvio = np.mean(dados), np.std(dados)
    densidade = (
        1 / (desvio * np.sqrt(2 * np.pi))
        * np.exp(-0.5 * ((xs - media) / desvio) ** 2)
    )
    ax.plot(xs, densidade, color="#253494", linewidth=2, label="Densidade normal")
    ax.axvline(sup_iqr, color=cor_outlier, linestyle="--", label="Limites IQR")
    ax.axvline(inf_iqr, color=cor_outlier, linestyle="--")
    ax.set_title("Distribuição dos valores")
    ax.set_xlabel("Valor")
    ax.set_ylabel("Densidade")
    ax.legend()

    # --- Painel 4: comparação IQR vs Z-score ---
    ax = axs[1, 1]
    metodos = ["IQR (1.5×)", "Z-score (3σ)"]
    n_out_iqr = int(np.sum((dados < inf_iqr) | (dados > sup_iqr)))
    n_out_z = int(np.sum((dados < inf_z) | (dados > sup_z)))
    barras = ax.bar(metodos, [n_out_iqr, n_out_z],
                    color=["#d7191c", "#fdae61"], edgecolor="black")
    ax.bar_label(barras, fmt="%d")
    ax.set_title("Nº de outliers detectados por método")
    ax.set_ylabel("Quantidade")

    plt.tight_layout(rect=[0, 0, 1, 0.96])
    plt.savefig(output, dpi=120, bbox_inches="tight")
    print(f"Gráfico salvo em: {output}")
    print(f"  Total de amostras : {len(dados)}")
    print(f"  Outliers (IQR)    : {n_out_iqr}")
    print(f"  Outliers (Z-score): {n_out_z}")
    print(f"  Limites IQR       : [{inf_iqr:.2f}, {sup_iqr:.2f}]")

    return fig


def main():
    parser = argparse.ArgumentParser(
        description="Gera um gráfico do comportamento de outliers."
    )
    parser.add_argument("--n", type=int, default=200,
                        help="Número de amostras (padrão: 200)")
    parser.add_argument("--seed", type=int, default=42,
                        help="Semente aleatória (padrão: 42)")
    parser.add_argument("--output", type=str,
                        default="outliers_comportamento.png",
                        help="Arquivo de saída do gráfico")
    args = parser.parse_args()

    dados = gerar_dados(n=args.n, seed=args.seed)
    plotar(dados, output=args.output)


if __name__ == "__main__":
    main()
