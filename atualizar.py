"""
Atualizacao mensal em um comando (Etapa 8).

Roda o motor (Etapa 5), auto-valida (reconciliacao) e gera o dashboard HTML
(Etapa 6) + as tabelas CSV/Excel. Rotina mensal:

    1. Coloque o novo Order_NN_2026.xlsx e o Received atualizado em data/
    2. python atualizar.py
    3. Confira a linha de reconciliacao (DIF deve ser 0/OK)
    4. Distribua saida/dashboard.html e saida/forecast_accuracy.xlsx

Leva menos de 10 minutos - o passo manual e so soltar o arquivo na pasta.
"""

import os
import sys

from forecast_accuracy import (Config, build_model, exporta_csv, exporta_excel,
                               exporta_json, reconcilia, yl)
from dashboard_build import build_payload, render


def main():
    outdir = "saida"
    cfg = Config(data_dir="data")
    print("Construindo o modelo a partir de data/ ...")
    model = build_model(cfg)
    os.makedirs(outdir, exist_ok=True)

    exporta_csv(model, outdir)
    exporta_excel(model, os.path.join(outdir, "forecast_accuracy.xlsx"))
    exporta_json(model, os.path.join(outdir, "dashboard_data.json"))
    render(build_payload(model), "dashboard_template.html",
           os.path.join(outdir, "dashboard.html"))

    a = model["agg"]["lag_honesto"]
    print(f"\nSafras: {[yl(s) for s in sorted(model['fc'].safra.unique())]}")
    print(f"Itens comparaveis: {len(model['overlap'])} | janela honesta: {[yl(t) for t in model['win_lag']]}")
    print(f"Headline lag-{cfg.lag_honesto}: bias {a['bias']*100:+.1f}% | "
          f"WAPE acum {a['wape_acum']*100:.1f}% | acuracia {a['acuracia']*100:.1f}%")
    ok = reconcilia(model)

    print(f"\nSaidas em {outdir}/: dashboard.html, forecast_accuracy.xlsx, "
          f"dashboard_data.json, *.csv")
    if not ok:
        print("ATENCAO: reconciliacao divergente - revise os arquivos de origem.")
        sys.exit(1)
    print("OK - reconciliacao com diferenca zero.")


if __name__ == "__main__":
    main()
