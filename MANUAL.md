# Manual — Acuracidade do Forecast de Importação (UPI)

Ferramenta de medição da acuracidade do forecast de compra intercompany (itens
comprados da matriz UPI) contra o recebimento físico. Motor em Python
(validável) + dashboard HTML autocontido + pasta Excel Power Pivot-ready.

---

## 1. Rotina mensal (< 10 minutos)

1. Copie o novo `Order_NN_2026.xlsx` (a safra do mês) e o `Received_itens_2026.xlsx`
   atualizado para a pasta **`data/`**. Não precisa renomear nada além do padrão
   `Order_NN_2026.xlsx` (NN = mês).
2. Rode um comando:
   ```bash
   python atualizar.py
   ```
3. Confira a linha **Reconciliação**: `DIF 0 OK`. Se aparecer `*** REVISAR`, algum
   arquivo de origem mudou de estrutura — pare e verifique antes de distribuir.
4. Distribua os dois entregáveis gerados em **`saida/`**:
   - `dashboard.html` — abre em qualquer navegador, offline (mande por e-mail/SharePoint);
   - `forecast_accuracy.xlsx` — tabelas para Power Pivot / análise no Excel.

O passo manual é só soltar o arquivo na pasta; o resto é automático (~7 s).

> Requisitos: Python 3.9+ com `pandas`, `openpyxl` e `pillow`
> (`pip install pandas openpyxl pillow`).

> **Logo:** se existir um `data/logo.png`, ele entra automaticamente nas abas de
> apresentação do Excel e no cabeçalho do dashboard HTML a cada geração.

---

## 2. Arquivos do projeto

| Arquivo | Papel |
|---|---|
| `forecast_accuracy.py` | Motor: ETL + modelo + métricas + exportações + reconciliação |
| `dashboard_build.py` | Monta o payload e injeta no template HTML |
| `dashboard_template.html` | Aparência/lógica do dashboard (JS puro, SVG, sem CDN) |
| `atualizar.py` | Rotina mensal em um comando (roda os dois acima + valida) |
| `atualizar.spec` · `build_exe.bat` · `BUILD_EXE.md` | Gerar um `.exe` para atualizar **sem Python** |
| `data/` | Entradas: `Order_NN_2026.xlsx` (safras) + `Received_itens_2026.xlsx` |
| `saida/` | Saídas: `dashboard.html`, `forecast_accuracy.xlsx`, CSVs, JSON |

Dados de origem e saídas com quantidades **não** são versionados (proprietários).

---

## 3. Premissas e decisões (Etapas 1–4)

- **Safra = número do arquivo** (`Order_NN` → mês N/2026), não a 1ª coluna —
  porque o Order_03 começa em Abril (Março já fechado no corte) e dataria colidindo
  com o Order_04. Ajustável em `SAFRAS`/leitura de `forecast_accuracy.py`.
- **Escopo = fornecedor `UPI`** (intercompany). Recebimentos de fornecedores locais
  ficam fora — o forecast da matriz não os cobre.
- **Recebido = `Qtd Delivered`** (entrada física), pelo mês do `Invoice Date`.
- **Chave do item = `COMPONENT UDB` normalizado** (sem sufixo `-BR`, sem traço final).
- **Leitura honesta = lag fixo N = 2** (lead time de importação `Invoice − PO`
  mediana ~2 meses = ponto em que a compra é travada).
- **Defasagem física sistemática = +1 mês** (o recebido chega ~1 mês depois do
  "mês de disponibilidade" previsto — trânsito/faturamento, não erro de previsão).
- **Mês corrente parcial fica fora** dos cálculos de acuracidade.
- **Defeitos tratados:** coluna corrompida do Order_06 (`01/05/20267` → Mai/27),
  coluna Fev/27 duplicada do Order_03 (mantém a 1ª), aba `DGR` do Order_04 ignorada.

---

## 4. Grão e lag

    Item × Mês-alvo × Safra (mês em que a previsão foi gerada)

**Lag = Mês-alvo − Safra** (em meses): Lag 0 = previsto no próprio mês, … até Lag 11.

### Dois modos de leitura

- **Lag fixo (honesto), N=2:** compara sempre a previsão travada 2 meses antes.
  Mede o forecast no ponto da decisão de compra. **Use para cobrar acurácia.**
- **Consolidado (gerencial):** compara com a previsão mais recente que cobre o mês.
  **Otimista** — usa revisões que não existiam na decisão. Só para leitura gerencial.

### Três leituras de timing (por causa da entrada física)

| Leitura | O que faz | Uso |
|---|---|---|
| (a) mês a mês | erro no grão item×mês | diagnóstico — superestima (pune timing) |
| (b) trimestral | agrega item×trimestre antes do erro | acompanhamento operacional |
| (c) **acumulada** | total do item na janela | **isola VOLUME** — a leitura mais justa |

**Recomendação:** para volume use **(c) acumulada**; (b) trimestral no operacional;
(a) mês a mês só como diagnóstico. Offset **+1 mês** disponível na leitura mensal.

---

## 5. Dicionário de métricas (Etapa 3)

Notação: `F` = previsto (pelo modo), `A` = recebido UPI, `e = A − F`.
Toda % agregada é **ponderada por volume** — nunca média simples.

| Métrica | Fórmula | Pergunta |
|---|---|---|
| WAPE (principal) | `Σ\|A−F\| / ΣA` | Quanto erramos? |
| Acurácia | `1 − WAPE` | Quão bons somos? |
| Bias / MPE | `Σ(A−F) / ΣF` | Erro é sistemático? p/ que lado? |
| Erro absoluto (un) | `Σ\|A−F\|` | Onde o erro pesa em volume/$ |
| MAPE | `média(\|A−F\|/A)`, só `A>0` e volume ≥ piso | Erro % típico de item relevante |
| Tracking signal | `Σ_corrente(e) / MAD` | Viés persistente? (\|TS\|>4 = fora de controle) |
| Forecast churn | `Σ\|F_v − F_{v−1}\| / F_última` p/ mesmo alvo | A previsão “balança” entre safras? |

### Classificação de padrão de erro (limiares aprovados)

| Classe | Regra | Ação |
|---|---|---|
| **Sob controle** | WAPE_acum ≤ 15% e WAPE_mês ≤ 35% | nenhuma |
| **Viés superprevisão** | bias ≤ −20% | corrigir método/premissa; revisar estoque (capital parado) |
| **Viés subprevisão** | bias ≥ +20% | revisar premissa + estoque de segurança (risco de ruptura) |
| **Erro de timing** | \|bias\| < 20%, WAPE_acum ≤ 25%, WAPE_mês ≥ 35% | tratar com lead time (+1 mês)/bucket, não mudar previsão |
| **Errático** | resto (dispersão que não colapsa no acumulado) | cobrir com estoque de segurança |
| **Não previsto** | recebido sem previsão na janela | incluir no forecast da matriz |

Parâmetros em `Config` (`forecast_accuracy.py`): `t_controle_acum=0.15`, `t_bias=0.20`,
`t_timing_acum=0.25`, `t_timing_mes=0.35`. Reestime conforme mais meses fecham.

---

## 6. Curva ABC, status e órfãos

- **ABC** por volume recebido: A ≤ 80%, B 80–95%, C 95–100%.
- **Status:** `novo` (entrou durante o ano), `descontinuado` (saiu das safras
  recentes), `ativo`. Novos/descontinuados são marcados e ficam fora de metas.
- **Órfãos** (nunca somem num join, ficam na aba Exceções): `previsto e não recebido`
  e `recebido sem previsão`. O ranking % usa piso de volume (curvas **A+B**).

---

## 7. Validação (Etapa 7)

A cada rodada, `atualizar.py` reconcilia **soma crua dos arquivos vs modelo** e
exige diferença zero (a única diferença admitida é a coluna duplicada do Order_03,
já descontada). Conferência pontual: abra a Ficha do item e compare com o arquivo
de origem — a matriz Safra×Mês-alvo mostra cada célula.

---

## 8. Carregar no Power Pivot (entrega 2)

No Excel: **Dados → Obter Dados → De um Arquivo → De Pasta de Trabalho**, aponte
para `saida/forecast_accuracy.xlsx`, carregue `grao_forecast` e `recebido_upi` como
tabelas-fato e relacione por `item`/`alvo`. A aba `mestre_itens` já traz as métricas
por item prontas para segmentar. A aba `leia-me` resume escopo e limiares.
