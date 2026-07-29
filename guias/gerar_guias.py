# -*- coding: utf-8 -*-
"""
Gera os 3 guias em PDF (didatico, executivo, tecnico) a partir do motor.

Os numeros NAO sao fixos: sao lidos de forecast_accuracy.build_model(), entao
basta rodar de novo depois de atualizar os dados em data/ que os guias saem
com os valores novos.

Uso:
    python guias/gerar_guias.py
Saida: saida/Guia_Acuracidade_Forecast.pdf, Guia_Executivo.pdf, Guia_Tecnico.pdf

Requisitos: pandas, openpyxl, pillow, playwright (+ 'playwright install chromium').
"""

import base64
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
from forecast_accuracy import Config, build_model, sel_consolidado  # noqa: E402

DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "saida")


# ---------- formatacao pt-BR ----------
def pc(x):  return "–" if x is None else f"{x*100:.1f}".replace(".", ",") + "%"
def pcs(x): return "–" if x is None else (("+" if x >= 0 else "") + f"{x*100:.1f}".replace(".", ",") + "%")
def mil(x): return f"{x:,.0f}".replace(",", ".")
def mi(x):  return f"{x/1e6:.2f}".replace(".", ",")


# ---------- numeros vivos, vindos do motor ----------
def figuras():
    m = build_model(Config(data_dir=DATA))
    v2 = m["views"]["2"]
    a = v2["agg"]
    mestre = m["mestre"]
    abc = mestre.abc.value_counts()
    cl = mestre.classe_label.value_counts()
    ok = mestre[mestre.A > 0]
    over = ok[ok.F_lag > ok.A]; under = ok[ok.A > ok.F_lag]
    asub = mestre[(mestre.abc == "A") & (mestre.classe_label == "Vies subprevisao")] \
        .sort_values("vol_recv", ascending=False)
    # offset de defasagem (consolidado, meses fechados)
    fF = sel_consolidado(m["fc_o"]).groupby("alvo").F.sum()
    aA = m["act_o"].groupby("alvo").recv.sum()
    Ts = list(m["meses_fechados"])
    off = {}
    for d in range(-2, 3):
        num = sum(abs(float(fF.get(t, 0)) - float(aA.get(t + d, 0))) for t in Ts)
        den = sum(float(aA.get(t + d, 0)) for t in Ts)
        off[d] = num / den * 100 if den else 0.0
    return dict(
        n=len(m["overlap"]), N=m["cfg"].lag_honesto, offset=m["cfg"].offset_defasagem,
        win=v2["window"], bias=a["bias"], wm=a["wape_mes"], wt=a["wape_trim"],
        wc=a["wape_acum"], acur=a["acuracia"], F=a["F"], A=a["A"],
        abcA=int(abc.get("A", 0)), abcB=int(abc.get("B", 0)), abcC=int(abc.get("C", 0)),
        cl={k: int(cl.get(k, 0)) for k in ["Erro de timing", "Vies subprevisao",
            "Vies superprevisao", "Nao previsto", "Sob controle", "Erratico"]},
        vol_over=float((over.F_lag - over.A).sum()), vol_under=float((under.A - under.F_lag).sum()),
        n_over=int(len(over)), n_under=int(len(under)),
        asub_n=int(len(asub)),
        asub=[(r["item"], str(r["desc"])[:26], r["vol_recv"], r["bias"], r["acuracia"])
              for _, r in asub.head(6).iterrows()],
        sens=[x for x in m["sensibilidade"] if x["lag"] <= 5],
        off=off,
    )


CSS = r"""
  @page { size:A4; margin:14mm 14mm 15mm 14mm; }
  *{box-sizing:border-box}
  body{font-family:-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#26221f;font-size:10.4pt;line-height:1.5;margin:0}
  h1,h2,h3{color:#a8410e;margin:0 0 6px}
  h2{font-size:14pt;border-bottom:2px solid #f0d3c0;padding-bottom:4px;margin:18px 0 6px}
  h3{color:#c2521a;font-size:11.5pt;margin:12px 0 4px}
  p{margin:6px 0} .muted{color:#6b645f}
  .hdr{display:flex;align-items:center;gap:16px;border-bottom:3px solid #d9601c;padding-bottom:10px;margin-bottom:6px}
  .hdr img{height:46px} .ht{font-size:18pt;font-weight:800;color:#d9601c;line-height:1.1} .hs{color:#6b645f;font-size:10pt;margin-top:2px}
  .cover{height:247mm;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;page-break-after:always}
  .cover img{height:74px;margin-bottom:26px}
  .cover .t{font-size:27pt;font-weight:800;color:#d9601c;line-height:1.15;max-width:16cm}
  .cover .s{font-size:13pt;color:#52514e;margin-top:14px}
  .cover .band{margin-top:30px;background:#fbf3ee;border:1px solid #f0d3c0;border-radius:12px;padding:14px 22px;color:#7a4a2c;font-size:10.5pt;max-width:14cm}
  .lead{background:#fbf3ee;border-left:4px solid #d9601c;padding:9px 13px;border-radius:0 8px 8px 0;margin:9px 0}
  .say{background:#eef7f3;border-left:4px solid #1b9e77;padding:8px 12px;border-radius:0 8px 8px 0;margin:8px 0}
  .say b{color:#12795a}
  .ex{background:#fdf6ee;border:1px dashed #e0a878;border-radius:8px;padding:8px 12px;margin:8px 0}
  .ana{color:#6b645f;font-style:italic;margin:6px 0}
  table{border-collapse:collapse;width:100%;margin:8px 0;font-size:9.6pt}
  th,td{border:1px solid #ecdccf;padding:5px 8px;text-align:left;vertical-align:top}
  th{background:#d9601c;color:#fff;font-weight:600}
  tr:nth-child(even) td{background:#fdf6f1}
  .kpirow{display:flex;gap:9px;margin:12px 0}
  .kpi{flex:1;border:1px solid #ecdccf;border-radius:10px;padding:11px 8px;text-align:center;background:#fffdfb}
  .kpi .v{font-size:19pt;font-weight:800;color:#c2521a;line-height:1}
  .kpi .v.red{color:#c0392b} .kpi .l{font-size:8pt;color:#6b645f;text-transform:uppercase;margin-top:5px}
  .metric{border:1px solid #ecdccf;border-radius:9px;padding:10px 13px;margin:9px 0;page-break-inside:avoid}
  .metric h3{margin-top:0}
  .pill{display:inline-block;background:#fce7da;color:#a8410e;border-radius:20px;padding:1px 9px;font-size:8.5pt;font-weight:700}
  .kbd{font-family:Consolas,monospace;background:#f3ebe4;border-radius:4px;padding:1px 5px}
  .formula{font-family:Consolas,'Courier New',monospace;background:#f7f2ec;border:1px solid #ecdccf;border-radius:6px;padding:6px 10px;margin:6px 0;font-size:9.3pt}
  .num{color:#c0392b;font-weight:700} .avoid{page-break-inside:avoid} ul{margin:6px 0 6px 18px} li{margin:4px 0}
"""


def _logo():
    p = os.path.join(DATA, "logo.png")
    if os.path.exists(p):
        return '<img src="data:image/png;base64,' + base64.b64encode(open(p, "rb").read()).decode() + '">'
    return ""


def _wrap(inner):
    return f'<!doctype html><html lang="pt-br"><head><meta charset="utf-8"><style>{CSS}</style></head><body>{inner}</body></html>'


def _hdr(t, s):
    return f'<div class="hdr">{_logo()}<div><div class="ht">{t}</div><div class="hs">{s}</div></div></div>'


# ---------- corpo dos 3 guias (numeros vindos de F) ----------
def html_didatico(F):
    win = ", ".join(F["win"])
    return _wrap(f"""
<div class="cover">{_logo()}
  <div class="t">Acuracidade do Forecast<br>de Importação</div>
  <div class="s">Guia para <b>entender</b> e <b>explicar</b> — passo a passo</div>
  <div class="band">Este material ensina, do zero e em linguagem simples, o que cada número
  significa: o que é <b>lag</b>, o que é <b>churn</b>, por que olhamos o erro de três jeitos e
  como apresentar tudo isso. Cada conceito vem com um exemplo real e uma frase pronta.</div>
</div>
<h2>Para que serve este guia</h2>
<p>Todo mês enviamos à matriz (UPI) uma <b>previsão</b> de quanto vamos precisar de cada item.
Depois o item <b>chega fisicamente</b>. A pergunta: <b>a previsão acertou?</b> — e, quando erra,
<b>para que lado</b>, <b>em quais itens</b> e <b>o que fazer</b>.</p>
<div class="lead"><b>A pegadinha:</b> o "recebido" é a <b>data de entrada física</b> (nota fiscal),
não a demanda. Um item previsto para abril pode chegar em maio só por transporte. Por isso o método
separa <b>erro de volume</b> de <b>erro de timing</b>.</div>

<h2>1. Safra, Mês-alvo e Lag</h2>
<table>
<tr><th>Palavra</th><th>O que é</th><th>Exemplo</th></tr>
<tr><td><b>Mês-alvo</b></td><td>o mês para o qual a previsão vale</td><td>Junho/2026</td></tr>
<tr><td><b>Safra</b></td><td>o mês em que a previsão nasceu</td><td>a previsão feita em Abril</td></tr>
<tr><td><b>Lag</b></td><td>a antecedência: <span class="kbd">Lag = mês-alvo − safra</span></td><td>Junho previsto em Abril = <b>Lag {F['N']}</b></td></tr>
</table>
<div class="ana">🔎 Analogia: como a previsão do tempo — prever o sábado na segunda (lag longo) erra
mais do que na sexta (lag curto).</div>
<div class="say">💡 <b>Frase pronta:</b> "Lag {F['N']} é a previsão que eu tinha na mão {F['N']} meses antes — a que usei para comprar."</div>

<h2>2. Os dois modos</h2>
<p><b>Lag fixo (honesto) — Lag {F['N']}:</b> compara sempre a previsão feita {F['N']} meses antes
(o lead time de importação), que foi a que <b>travou a compra</b>. <b>Consolidado:</b> compara com a
previsão mais recente — otimista, só para leitura gerencial.</p>
<div class="say">💡 <b>No dashboard</b>, o seletor "Leitura (lag)" alterna Lag 0 (mesmo mês) → Consolidado e tudo se recalcula.</div>

<h2>3. Três leituras do erro</h2>
<table>
<tr><th>Leitura</th><th>O que faz</th><th>Serve para</th></tr>
<tr><td>(a) Mês a mês</td><td>erro em cada mês</td><td>diagnóstico — injusta (pune o timing)</td></tr>
<tr><td>(b) Trimestral</td><td>soma o trimestre</td><td>operacional (absorve ±1 mês)</td></tr>
<tr><td>(c) Acumulada</td><td>soma tudo na janela</td><td><b>isola o VOLUME</b> — a mais justa</td></tr>
</table>
<div class="ex">📌 <b>Hoje (Lag {F['N']}, {win}):</b> erro mês a mês <span class="num">{pc(F['wm'])}</span>,
mas acumulado <span class="num">{pc(F['wc'])}</span> — grande parte do "erro" mensal é timing (chega +{F['offset']} mês depois).</div>
<div class="say">💡 <b>Frase:</b> "No acumulado, o erro real de volume é {pc(F['wc'])}."</div>

<h2>4. As métricas</h2>
<div class="metric"><h3>Bias (viés) <span class="pill">para que lado</span></h3>
<p><span class="kbd">(Recebido − Previsto) ÷ Previsto</span>. Hoje: <span class="num">{pcs(F['bias'])}</span>
— {'subestimamos' if F['bias']>=0 else 'superestimamos'} de forma sistemática (corrigível).</p></div>
<div class="metric"><h3>WAPE / Acurácia <span class="pill">quanto erro</span></h3>
<p><span class="kbd">WAPE = Σ|A−F| ÷ ΣA</span> (ponderado por volume). <b>Acurácia = 1 − WAPE = {pc(F['acur'])}</b>.</p></div>
<div class="metric"><h3>Tracking signal <span class="pill">alarme de viés</span></h3>
<p><span class="kbd">Σ erros ÷ erro médio</span>. <b>|TS| &gt; 4</b> = viés travado, precisa de ação.</p></div>
<div class="metric"><h3>Churn <span class="pill">a previsão balança?</span></h3>
<p>Quanto a previsão de um mesmo alvo muda entre safras. Alto churn = a matriz não consegue planejar.</p></div>

<h2>5. Os 6 padrões de erro</h2>
<table>
<tr><th>Padrão</th><th>Itens</th><th>Ação</th></tr>
<tr><td>Sob controle</td><td>{F['cl']['Sob controle']}</td><td>nada</td></tr>
<tr><td>Viés subprevisão</td><td>{F['cl']['Vies subprevisao']}</td><td>corrigir premissa + estoque (ruptura)</td></tr>
<tr><td>Viés superprevisão</td><td>{F['cl']['Vies superprevisao']}</td><td>corrigir método (capital parado)</td></tr>
<tr><td>Erro de timing</td><td>{F['cl']['Erro de timing']}</td><td>lead time/bucket — não mexer na previsão</td></tr>
<tr><td>Errático</td><td>{F['cl']['Erratico']}</td><td>estoque de segurança</td></tr>
<tr><td>Não previsto</td><td>{F['cl']['Nao previsto']}</td><td>incluir no forecast</td></tr>
</table>

<h2>6. Curva ABC</h2>
<p><b>A = {F['abcA']} itens</b> (80% do volume) · B = {F['abcB']} · C = {F['abcC']}. Arrume os A primeiro.</p>

<h2>7. A história em 5 frases</h2>
<div class="lead avoid">
<p>1. Comparamos a previsão que travou a compra ({F['N']} meses antes) com o que entrou.</p>
<p>2. Metade do erro mensal é o navio chegando 1 mês depois — no acumulado, o volume erra {pc(F['wc'])}.</p>
<p>3. {'Subestimamos' if F['bias']>=0 else 'Superestimamos'} ~{pc(abs(F['bias']))}, concentrado nos {F['abcA']} itens que são 80% do volume.</p>
<p>4. Cada item tem um padrão: viés, timing ou errático — e cada um pede uma ação.</p>
<p>5. Meta: subir a acurácia de volume ({pc(F['acur'])}) corrigindo o viés dos itens A.</p>
</div>
""")


def html_exec(F):
    asub = ", ".join(x[0] for x in F["asub"][:4])
    return _wrap(_hdr("Acuracidade do Forecast — Sumário Executivo",
                      f"UPI (intercompany) · PCP · leitura honesta (lag {F['N']})") + f"""
<h2>O essencial</h2>
<div class="kpirow">
  <div class="kpi"><div class="v">{pc(F['acur'])}</div><div class="l">Acurácia de volume</div></div>
  <div class="kpi"><div class="v red">{pcs(F['bias'])}</div><div class="l">Bias — {'subprevisão' if F['bias']>=0 else 'superprevisão'}</div></div>
  <div class="kpi"><div class="v">{pc(F['wc'])}</div><div class="l">Erro de volume (WAPE)</div></div>
  <div class="kpi"><div class="v">{F['abcA']} itens</div><div class="l">= 80% do volume (curva A)</div></div>
</div>
<div class="lead">Comparamos a previsão que <b>travou a compra</b> (~{F['N']} meses antes) com o que
realmente entrou. Parte do "erro" mensal é o <b>navio chegando ~{F['offset']} mês depois</b>; no
<b>volume acumulado</b>, o erro real é <b>{pc(F['wc'])}</b> — e é de
<b>{'subprevisão' if F['bias']>=0 else 'superprevisão'}</b>.</div>
<div class="say">No <b>dashboard</b>, o seletor "Leitura (lag)" alterna entre Lag 0 (mesmo mês),
Lag {F['N']} (honesto) e Consolidado, recalculando tudo na hora.</div>

<h2>O que isso significa para o negócio</h2>
<table>
<tr><th>Fato</th><th>Impacto</th></tr>
<tr><td>Recebemos <b>{mi(F['vol_under'])} milhão de un a mais</b> do que previmos (vs {mi(F['vol_over'])} M a mais previstas)</td><td><b>Risco de ruptura</b> e frete aéreo</td></tr>
<tr><td><b>{F['asub_n']} itens da curva A</b> têm subprevisão crônica</td><td>O erro dói onde mais custa</td></tr>
<tr><td>Boa parte do erro mensal é <b>timing de importação</b>, não previsão</td><td>Não se resolve mudando o forecast</td></tr>
</table>

<h2>Onde agir — prioridades</h2>
<table>
<tr><th>#</th><th>Ação</th><th>Efeito</th></tr>
<tr><td>1</td><td>Corrigir o <b>viés dos itens A</b> ({asub}…)</td><td>menos ruptura; acurácia sobe</td></tr>
<tr><td>2</td><td><b>Estoque de segurança por classe de acurácia</b></td><td>protege sem inflar o total</td></tr>
<tr><td>3</td><td>Itens de <b>timing</b>: lead time (+{F['offset']} mês), não mexer na previsão</td><td>evita ordens indevidas</td></tr>
<tr><td>4</td><td><b>Spares</b> (churn alto): reposição por ponto de pedido</td><td>forecast estável p/ a matriz</td></tr>
<tr><td>5</td><td><b>Pauta com a matriz</b>: viés e freeze de {F['N']} meses</td><td>planejamento alinhado</td></tr>
</table>

<h2>Meta</h2>
<div class="lead avoid">Subir a <b>acurácia de volume</b> (hoje {pc(F['acur'])}) corrigindo o viés dos itens A,
e reduzir rupturas nos itens de alto giro — sem inflar o estoque total.</div>
""")


def html_tec(F):
    sens = "".join(
        f"<tr><td>{'<b>'+str(x['lag'])+'</b>' if x['lag']==F['N'] else x['lag']}</td>"
        f"<td>{pcs(x['bias'])}</td><td>{pc(x['wape_acum'])}</td></tr>" for x in F["sens"])
    offh = "".join(f"<th>{d:+d}</th>" for d in range(-2, 3))
    offv = "".join(
        (f'<td class="num">' if d == F["offset"] else "<td>") + f"{F['off'][d]:.1f}".replace(".", ",") + "%</td>"
        for d in range(-2, 3))
    asub = "".join(f"<tr><td>{it}</td><td>{d}</td><td>{mil(v)}</td><td>{pcs(b)}</td><td>{pc(a)}</td></tr>"
                   for it, d, v, b, a in F["asub"])
    return _wrap(_hdr("Acuracidade do Forecast — Nota Técnica",
                      "Metodologia, fórmulas, limiares e validação") + f"""
<h2>1. Modelo de dados e escopo</h2>
<ul>
<li><b>Grão:</b> Item × Mês-alvo × Safra; <b>safra = número do arquivo</b> (Order_03 começa em Abril).</li>
<li><b>Escopo:</b> fornecedor <span class="kbd">UPI</span>; <span class="kbd">Qtd Delivered</span>; item = <span class="kbd">COMPONENT UDB</span> normalizado.</li>
<li><b>Defeitos tratados:</b> coluna corrompida do Order_06 → Mai/27; coluna Fev/27 duplicada do Order_03; DGR do Order_04 ignorada.</li>
<li><b>Cobertura:</b> {F['n']} itens comparáveis · janela honesta (lag {F['N']}): {', '.join(F['win'])}.</li>
</ul>

<h2>2. Sensibilidade por lag (janela fechada)</h2>
<table><tr><th>Lag</th><th>Bias</th><th>WAPE acum (volume)</th></tr>{sens}</table>
<p>Lead time de importação ~{F['N']} meses → o buy é travado no <b>lag {F['N']}</b> (última previsão acionável).</p>

<h2>3. Defasagem sistemática (offset)</h2>
<p>Alinhando <span class="kbd">Previsto[T]</span> com <span class="kbd">Recebido[T+d]</span>:</p>
<table><tr><th>Offset d</th>{offh}</tr><tr><td>WAPE mensal</td>{offv}</tr></table>
<p>Mínimo em <b>d = +{F['offset']}</b>: o recebido chega ~{F['offset']} mês depois do previsto (transporte, não erro).</p>

<h2>4. Métricas — fórmulas</h2>
<div class="formula">Bias = Σ(A−F)/ΣF    WAPE = Σ|A−F|/ΣA    Acurácia = 1 − WAPE</div>
<div class="formula">WAPE: (a) item×mês  (b) item×trimestre  (c) item na janela (volume)</div>
<div class="formula">Tracking signal = Σ_corrente(e)/MAD ; |TS|&gt;4 fora de controle   ·   Churn = Σ|F_v−F_(v-1)|/F_última</div>
<p><b>Resultado (lag {F['N']}):</b> ΣF {mil(F['F'])} · ΣA {mil(F['A'])} · bias {pcs(F['bias'])} ·
WAPE mês {pc(F['wm'])} · trim {pc(F['wt'])} · acum <b>{pc(F['wc'])}</b> · acurácia {pc(F['acur'])}.</p>

<h2>5. Classificação (limiares 15/20/25/35%)</h2>
<p>Distribuição ({F['n']} itens): timing {F['cl']['Erro de timing']} · subprevisão {F['cl']['Vies subprevisao']} ·
superprevisão {F['cl']['Vies superprevisao']} · não previsto {F['cl']['Nao previsto']} ·
sob controle {F['cl']['Sob controle']} · errático {F['cl']['Erratico']}. ABC: A {F['abcA']} · B {F['abcB']} · C {F['abcC']}.</p>
<h3>Itens ABC A com subprevisão (prioridade)</h3>
<table><tr><th>Item</th><th>Descrição</th><th>Vol</th><th>Bias</th><th>Acur</th></tr>{asub}</table>

<h2>6. Validação e leitura por lag</h2>
<p>A cada execução, <span class="kbd">reconcilia()</span> confere a soma crua vs modelo (diferença zero).
O dashboard expõe um seletor de Lag (0..4 + consolidado) e o Excel traz a aba
<span class="kbd">metricas_por_lag</span> (tidy) para o mesmo corte via Segmentação de Dados.</p>
""")


def _launch(p):
    cand = os.environ.get("CHROME_PATH")
    pinned = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
    if cand and os.path.exists(cand):
        return p.chromium.launch(executable_path=cand, args=["--no-sandbox"])
    if os.path.exists(pinned):
        return p.chromium.launch(executable_path=pinned, args=["--no-sandbox"])
    return p.chromium.launch(args=["--no-sandbox"])


def main():
    from playwright.sync_api import sync_playwright
    F = figuras()
    os.makedirs(OUT, exist_ok=True)
    jobs = [("Guia_Acuracidade_Forecast", html_didatico(F)),
            ("Guia_Executivo", html_exec(F)),
            ("Guia_Tecnico", html_tec(F))]
    with sync_playwright() as p:
        b = _launch(p)
        for nome, html in jobs:
            tmp = os.path.join(OUT, nome + ".html")
            open(tmp, "w", encoding="utf-8").write(html)
            pg = b.new_page(); pg.goto("file://" + tmp); pg.wait_for_timeout(350)
            pg.pdf(path=os.path.join(OUT, nome + ".pdf"), format="A4", print_background=True,
                   margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})
            pg.close(); print("gerado:", os.path.join(OUT, nome + ".pdf"))
        b.close()


if __name__ == "__main__":
    main()
