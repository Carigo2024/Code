"""
Atualizacao mensal do dashboard + Excel (Python OU .exe empacotado).

Rotina: coloque os Order_NN_2026.xlsx, o Received_itens_2026.xlsx e (opcional)
um logo.png numa pasta 'data' ao lado deste programa; execute; pegue os
resultados em 'saida' (dashboard.html + forecast_accuracy.xlsx).

- Como script:     python atualizar.py
- Como executavel: duplo-clique em AtualizarDashboard.exe
  (ver BUILD_EXE.md para gerar o .exe uma unica vez)
"""

import os
import sys

from forecast_accuracy import (Config, build_model, exporta_csv, exporta_excel,
                               exporta_json, reconcilia, yl)
from dashboard_build import build_payload, render


def _app_dir():
    """Pasta onde o programa roda: do .exe quando empacotado, senao do script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _resource(nome):
    """Arquivos empacotados no .exe (ex.: o template) ficam em sys._MEIPASS."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nome)


def executar():
    app = _app_dir()
    data_dir = os.path.join(app, "data")
    outdir = os.path.join(app, "saida")
    template = _resource("dashboard_template.html")

    if not os.path.isdir(data_dir):
        raise SystemExit(
            f"Pasta 'data' nao encontrada em:\n  {app}\n"
            "Crie a pasta 'data' e coloque os Order_NN_2026.xlsx, o "
            "Received_itens_2026.xlsx e (opcional) o logo.png dentro dela.")

    cfg = Config(data_dir=data_dir, logo_path=os.path.join(data_dir, "logo.png"))
    print(f"Lendo arquivos de: {data_dir}")
    model = build_model(cfg)
    os.makedirs(outdir, exist_ok=True)

    exporta_csv(model, outdir)
    exporta_excel(model, os.path.join(outdir, "forecast_accuracy.xlsx"))
    exporta_json(model, os.path.join(outdir, "dashboard_data.json"))
    render(build_payload(model), template, os.path.join(outdir, "dashboard.html"))

    a = model["agg"]["lag_honesto"]
    print(f"\nSafras: {[yl(s) for s in sorted(model['fc'].safra.unique())]}")
    print(f"Itens comparaveis: {len(model['overlap'])} | "
          f"janela honesta: {[yl(t) for t in model['win_lag']]}")
    print(f"Headline lag-{cfg.lag_honesto}: bias {a['bias']*100:+.1f}% | "
          f"WAPE acum {a['wape_acum']*100:.1f}% | acuracia {a['acuracia']*100:.1f}%")
    ok = reconcilia(model)
    print(f"\nSaidas em: {outdir}")
    print("  - dashboard.html  (abrir no navegador)")
    print("  - forecast_accuracy.xlsx  (Excel)")
    if not ok:
        print("\nATENCAO: reconciliacao divergente - revise os arquivos de origem.")
        return 1
    print("\nOK - reconciliacao com diferenca zero.")
    return 0


def main():
    codigo = 1
    try:
        codigo = executar()
    except SystemExit as e:
        print(e)
    except Exception as e:  # noqa: BLE001 - queremos mostrar qualquer erro ao usuario
        import traceback
        traceback.print_exc()
        print(f"\nERRO: {e}")
    # no .exe (duplo-clique), pausa para o usuario ler a saida antes de fechar
    if getattr(sys, "frozen", False):
        try:
            input("\nPressione Enter para fechar...")
        except EOFError:
            pass
    sys.exit(codigo)


if __name__ == "__main__":
    main()
