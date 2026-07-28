"""
Gera o dashboard HTML autocontido (Etapa 6) a partir do motor da Etapa 5.

Le os dados via forecast_accuracy.build_model(), enriquece com o detalhe por
item (matriz Safra x Mes-alvo, acuracidade por lag, serie previsto x recebido)
e injeta tudo inline num template HTML - sem servidor, sem CDN, abre offline
em qualquer navegador.

Uso:
    python dashboard_build.py                    # -> saida/dashboard.html
    python dashboard_build.py --outfile x.html
"""

import argparse
import json
import os

import numpy as np
import pandas as pd

from forecast_accuracy import (Config, build_model, yl, sel_lag, sel_consolidado,
                               malha, wape)


def _lag_accuracy_item(fc_it, act_it, closed_max):
    """Acuracidade por lag do proprio item (meses fechados)."""
    out = []
    for L in sorted(set(fc_it.lag)):
        if L < 0 or L > 11:
            continue
        s = fc_it[fc_it.lag == L][["alvo", "fcst"]]
        w = sorted(t for t in set(s.alvo) & set(act_it.alvo) if t <= closed_max)
        if not w:
            continue
        F = s[s.alvo.isin(w)].fcst.sum()
        A = act_it[act_it.alvo.isin(w)].recv.sum()
        if A <= 0 and F <= 0:
            continue
        wc = abs(A - F) / A if A > 0 else None
        out.append(dict(lag=int(L), F=float(F), A=float(A),
                        wape_acum=(None if wc is None else round(wc, 3))))
    return out


def _matriz_item(fc_it, act_it, tmax):
    """Matriz Safra(linhas) x Mes-alvo(cols) com fcst + linha de realizado."""
    targets = sorted(t for t in set(fc_it.alvo) if t <= tmax)
    vintages = sorted(set(fc_it.safra))
    piv = (fc_it[fc_it.alvo <= tmax]
           .pivot_table(index="safra", columns="alvo", values="fcst", aggfunc="first"))
    matrix = []
    for v in vintages:
        row = []
        for t in targets:
            val = piv.loc[v, t] if (v in piv.index and t in piv.columns) else np.nan
            row.append(None if pd.isna(val) else round(float(val), 1))
        matrix.append(row)
    actual = []
    amap = dict(zip(act_it.alvo, act_it.recv))
    for t in targets:
        actual.append(round(float(amap[t]), 1) if t in amap else None)
    return dict(vintages=[yl(v) for v in vintages],
                targets=[yl(t) for t in targets], matrix=matrix, actual=actual)


def build_payload(model):
    cfg = model["cfg"]
    fc_o, act_o = model["fc_o"], model["act_o"]
    closed_max = model["closed_max"]
    tmax = closed_max + 3  # matriz mostra horizonte proximo

    # ---- visao geral: series mensais (meses fechados) ----
    meses = model["meses_fechados"]
    A_m = act_o.groupby("alvo").recv.sum()
    Fl = sel_lag(fc_o, cfg.lag_honesto).groupby("alvo").F.sum()
    Fc = sel_consolidado(fc_o).groupby("alvo").F.sum()
    overview = dict(
        meses=[yl(t) for t in meses],
        A=[round(float(A_m.get(t, 0)), 1) for t in meses],
        F_lag=[None if t not in Fl.index else round(float(Fl.get(t)), 1) for t in meses],
        F_cons=[round(float(Fc.get(t, 0)), 1) for t in meses],
    )
    # erro % mensal (consolidado) para sparkline/tendencia
    err = [((A_m.get(t, 0) - Fc.get(t, 0)) / Fc.get(t, 1) * 100) if Fc.get(t, 0) else None
           for t in meses]
    overview["err_pct"] = [None if e is None else round(e, 1) for e in err]
    # tendencia conservadora: media |erro%| da 1a metade vs 2a metade dos meses
    vals = [abs(e) for e in err if e is not None]
    if len(vals) >= 4:
        h = len(vals) // 2
        m1, m2 = float(np.mean(vals[:h])), float(np.mean(vals[h:]))
        if m2 < m1 * 0.8:
            overview["tendencia"] = "melhorando"
        elif m2 > m1 * 1.2:
            overview["tendencia"] = "piorando"
        else:
            overview["tendencia"] = "estavel"
    else:
        overview["tendencia"] = "n/d"
    overview["tendencia_n"] = len(vals)

    # ---- detalhe por item ----
    detalhe = {}
    for it in model["overlap"]:
        fi = fc_o[fc_o.item == it]
        ai = act_o[act_o.item == it]
        detalhe[it] = dict(
            matriz=_matriz_item(fi, ai, tmax),
            lag_acc=_lag_accuracy_item(fi, ai, closed_max),
        )

    payload = dict(
        meta=dict(escopo=cfg.fornecedor, lag=cfg.lag_honesto, offset=cfg.offset_defasagem,
                  mes_parcial=cfg.mes_corrente_parcial, n_itens=len(model["overlap"]),
                  win_lag=[yl(t) for t in model["win_lag"]],
                  win_full=[yl(t) for t in model["win_full"]],
                  limiares=dict(controle_acum=cfg.t_controle_acum, controle_mes=cfg.t_controle_mes,
                                bias=cfg.t_bias, timing_acum=cfg.t_timing_acum,
                                timing_mes=cfg.t_timing_mes)),
        agg=model["agg"],
        sensibilidade=model["sensibilidade"],
        cobertura=model["cobertura"],
        overview=overview,
        mestre=model["mestre"].where(pd.notnull(model["mestre"]), None).to_dict("records"),
        series=model["series"],
        detalhe=detalhe,
        rankings={k: v.where(pd.notnull(v), None).assign(
                     erro_abs_un=(v.A - v.F_lag).abs().round(0)
                  )[["item", "desc", "familia", "abc", "F_lag", "A", "bias",
                     "wape_acum", "erro_abs_un"]].to_dict("records")
                  for k, v in model["rankings"].items()},
        orfaos=dict(
            previsto_nao_recebido=model["orfaos"]["previsto_nao_recebido"][:300],
            recebido_sem_previsao=model["orfaos"]["recebido_sem_previsao"][:300],
        ),
    )
    return payload


def render(payload, template_path, outfile):
    with open(template_path, encoding="utf-8") as f:
        tpl = f.read()
    data_js = json.dumps(payload, ensure_ascii=False, default=float)
    html = tpl.replace("/*__DATA__*/", data_js)
    os.makedirs(os.path.dirname(outfile) or ".", exist_ok=True)
    with open(outfile, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Dashboard gerado: {outfile}  ({len(html)//1024} KB)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir", default="data")
    ap.add_argument("--template", default="dashboard_template.html")
    ap.add_argument("--outfile", default="saida/dashboard.html")
    args = ap.parse_args()
    model = build_model(Config(data_dir=args.data_dir))
    payload = build_payload(model)
    render(payload, args.template, args.outfile)


if __name__ == "__main__":
    main()
