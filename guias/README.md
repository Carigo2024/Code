# Guias em PDF (didático · executivo · técnico)

`gerar_guias.py` produz os três PDFs a partir do **motor** — os números não são
fixos: são lidos de `forecast_accuracy.build_model()`. Atualize os dados em
`data/`, rode, e os guias saem com os valores novos.

```bash
# uma vez, na máquina que vai gerar os PDFs:
pip install pandas openpyxl pillow playwright
playwright install chromium

# gerar (a partir da raiz do projeto):
python guias/gerar_guias.py
```

Saída em `saida/`:
- `Guia_Acuracidade_Forecast.pdf` — didático (aprender/ensinar)
- `Guia_Executivo.pdf` — 1 página, diretoria (KPIs, impacto, prioridades)
- `Guia_Tecnico.pdf` — nota técnica (fórmulas, sensibilidade, offset, validação)

O logo (`data/logo.png`) entra automaticamente. Se o Chromium não for encontrado,
aponte `CHROME_PATH` para o executável do navegador.
