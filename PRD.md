# PRD — Sistema de Acuracidade de Forecast de Importação (UPI)

**Produto:** ferramenta de medição, diagnóstico e melhoria da acuracidade do
forecast de compra intercompany (itens comprados da matriz UPI).
**Contexto:** indústria odontológica/farmacêutica (Ultradent), PCP, Brasil, ambiente
regulado (ANVISA/BPF), ERP Protheus, análises usuais em Excel pt-BR (Power
Query/Power Pivot). **Documento:** engenharia reversa do projeto entregue.

---

## 1. Problema

Todo mês a subsidiária envia à matriz (UPI) um forecast de compra com 12 meses
rolantes. O mesmo mês-alvo é previsto várias vezes (uma por safra). Não havia uma
medição defensável de "a previsão acertou?", e três armadilhas atrapalhavam:

1. **Vintage/lag implícitos** — sem escolher *qual* previsão comparar, qualquer
   número de acurácia é arbitrário.
2. **"Recebido" ≠ demanda** — o recebido é data de entrada física (nota fiscal),
   sujeita a lead time de importação; comparar mês a mês acusa erro de previsão
   onde há apenas defasagem de embarque.
3. **Métrica mal definida** — média simples de percentuais, MAPE com zeros, e
   confusão entre erro de volume e erro de timing.

Consequência de negócio: erro para baixo (subprevisão) → ruptura e frete aéreo;
erro para cima (superprevisão) → capital parado, ocupação de armazém, risco de
vencimento (produtos com shelf life). Forecast instável destrói a credibilidade
junto à matriz e piora o atendimento.

## 2. Objetivos

- **O1.** Medir a acurácia com método defensável, separando **erro de volume** de
  **erro de timing**.
- **O2.** Localizar onde o erro se concentra (item, família, curva ABC) e para que
  lado (viés).
- **O3.** Permitir análise **item a item** de qualquer item do portfólio.
- **O4.** Oferecer uma **rotina mensal de baixo esforço** (<10 min).
- **O5.** Traduzir os números em **ações** de redução de erro.

**Métrica de sucesso (baseline atual):** acurácia de volume (lag 2) ≈ 78%; meta de
evolução para ~85% corrigindo o viés dos itens da curva A.

## 3. Personas / usuários

| Persona | Uso |
|---|---|
| **Especialista de PCP** (owner) | rotina mensal, ficha do item, tabela mestre, decisões de compra/estoque |
| **Gestão / diretoria** | sumário executivo: KPIs, risco, prioridades |
| **Analista / time** | dicionário de métricas, validação, calibração de limiares |
| **Matriz (UPI)** | pauta de alinhamento (viés, freeze, estabilidade) |

## 4. Escopo

**Dentro:** ingestão dos arquivos crus; modelo Item × Mês-alvo × Safra; dois modos
de leitura; três leituras de timing; dicionário de métricas; classificação de
padrão de erro; dashboard interativo; pasta Excel; versão nativa Excel (sem
Python); executável; guias em PDF; rotina mensal; validação/reconciliação; plano
de ação.

**Fora:** previsão em si (o produto mede, não gera forecast); integração online com
ERP; multiusuário/hospedagem corporativa; leitura em valor financeiro com câmbio
(camada opcional futura); itens de fornecedores locais (fora do escopo do forecast
da matriz).

## 5. Premissas e decisões metodológicas

- **Safra = número do arquivo** (`Order_NN` → mês N/2026), não a 1ª coluna — o
  Order_03 começa em Abril e, datado pela coluna, colidiria com o Order_04.
- **Escopo do recebido = fornecedor `UPI`** (intercompany); quantidade =
  `Qtd Delivered`; mês pelo `Invoice Date`.
- **Chave do item = `COMPONENT UDB` normalizado** (remove sufixo `-BR` e traço final).
- **Leitura honesta = lag fixo N = 2**, derivado do lead time de importação real
  (`Invoice − PO`, mediana ~2 meses): é a última previsão acionável antes do commit.
- **Defasagem física sistemática = +1 mês** (recebido chega depois do previsto).
- **Mês corrente parcial** fica fora dos cálculos de acurácia.
- **Defeitos dos arquivos tratados:** coluna corrompida do Order_06
  (`01/05/20267` → Mai/27); coluna Fev/27 duplicada do Order_03 (mantém a 1ª); aba
  DGR do Order_04 ignorada.

## 6. Modelo de dados

**Grão:** `Item × Mês-alvo × Safra`. **Lag = mês-alvo − safra** (0 a 11).
**Tabelas-fato:** `grao_forecast` (item, alvo, safra, lag, fcst) e `recebido_upi`
(item, alvo, recv). **Dimensão de item:** descrição, família, marca, UM, ABC,
status (novo/descontinuado/ativo), churn.

## 7. Requisitos funcionais

**Motor / cálculo**
- RF1. Ingestão parametrizada de qualquer `Order_NN` (adicionar um mês novo não
  exige refazer nada) + `Received`, tratando os defeitos acima.
- RF2. Dois **modos de leitura**: **lag fixo N** (honesto) e **consolidado** (última
  safra, gerencial/otimista).
- RF3. Três **leituras de timing**: (a) mês a mês, (b) trimestral, (c) acumulada;
  quantificar o **offset** de defasagem e recomendar a leitura mais justa
  (acumulada para volume).
- RF4. **Métricas** (dicionário §9), sempre ponderadas por volume; nunca média
  simples de percentuais.
- RF5. **Classificação de padrão de erro** por item (6 classes) com limiares
  aprovados e **ação** associada.
- RF6. **Curva ABC** por volume; **status** de item; **órfãos** (previsto-não-recebido
  e recebido-sem-previsão) em tela própria; itens novos/descontinuados marcados.
- RF7. **Rankings**: 4 (variação absoluta em unidades e variação % com piso de
  volume A+B), nas duas direções (super/sub).
- RF8. **Views por lag**: métricas pré-calculadas por lag (0..4 + consolidado) para
  seleção interativa.
- RF9. **Reconciliação** automática (soma crua vs modelo = diferença zero, com a
  coluna duplicada explicada); falha explícita se divergir.

**Dashboard (HTML autocontido)**
- RF10. **Seletor de Lag global** (Lag 0 = mesmo mês … Lag 4, + Consolidado) que
  reprocessa visão geral, rankings, tabela mestre, ficha e exceções.
- RF11. Telas: (1) Visão geral (KPIs + Previsto×Recebido mensal), (2) Maiores
  variações (4 rankings), (3) **Ficha do item** — qualquer item, série mês a mês,
  matriz Safra×Mês-alvo com realizado sobreposto, acurácia por lag do próprio item,
  padrão + ação, (3b) Tabela mestre ordenável/filtrável/exportável, (4) Acurácia
  por lag, (5) Estabilidade (churn), (6) Exceções.
- RF12. Filtros: lag, curva ABC, família, busca de item.

**Excel / Power Pivot**
- RF13. Pasta `forecast_accuracy.xlsx`: `leia-me`, `mestre_itens` (com formatação de
  %), `metricas_por_lag` (tidy, para Segmentação por lag), `cobertura`, tabelas-fato.
- RF14. **Versão nativa** (Power Query M + DAX) que lê os arquivos crus e atualiza
  com "Atualizar Tudo", **sem Python**.

**Distribuição / operação**
- RF15. **Rotina mensal em um comando** (`atualizar.py`) que gera dashboard + Excel +
  CSV/JSON e valida.
- RF16. **Executável** (`.exe` via PyInstaller) para atualizar sem Python instalado.
- RF17. **Guias em PDF** (didático, executivo, técnico) gerados a partir do motor
  (números vivos).
- RF18. **Branding**: paleta laranja e logo da empresa (`data/logo.png`) embutidos
  automaticamente em Excel, dashboard e PDFs.

## 8. Requisitos não-funcionais

- RNF1. **Autocontido / offline**: dashboard em um arquivo, sem servidor nem CDN
  (adequado a ambiente regulado sem hosting externo).
- RNF2. **Baixo esforço mensal**: atualização em <10 min (um comando / um clique).
- RNF3. **Reprodutível e auditável**: cálculo em Python testável; reconciliação a
  cada rodada.
- RNF4. **Acessibilidade visual**: séries de gráfico distinguíveis e CVD-safe
  (Previsto = teal, Recebido = laranja).
- RNF5. **Privacidade**: dados de origem e saídas com quantidades **não** versionados;
  só o código vai para o repositório.
- RNF6. **Parametrização**: N do lag, limiares de classe (15/20/25/35%), cortes ABC
  (80/95), offset — todos em `Config`.

## 9. Dicionário de métricas

| Métrica | Fórmula | Nível | Exceção | Pergunta |
|---|---|---|---|---|
| WAPE (principal) | `Σ\|A−F\|/ΣA` | item/família/geral, 3 leituras | ΣA=0 → n/d | Quanto erramos? |
| Acurácia | `1−WAPE` | idem | idem | Quão bons somos? |
| Bias / MPE | `Σ(A−F)/ΣF` | idem | ΣF=0 → "não previsto" | Erro sistemático? p/ que lado? |
| Erro absoluto (un) | `Σ\|A−F\|` | item/família | — | Onde pesa em volume/$ |
| MAPE | `média(\|A−F\|/A)` | item | só A>0 e volume≥piso | Erro % típico de item relevante |
| Tracking signal | `Σ_corrente(e)/MAD` | item | \|TS\|>4 fora de controle | Viés persistente? |
| Forecast churn | `Σ\|F_v−F_(v−1)\|/F_última` | item/geral | ≥2 safras cobrindo T | A previsão balança? |

**Classificação (árvore, 1ª regra que casa):** Sob controle (WAPE_acum≤15% e
WAPE_mês≤35%) · Viés superprevisão (bias≤−20%) · Viés subprevisão (bias≥+20%) ·
Erro de timing (|bias|<20% e WAPE_acum≤25% e WAPE_mês≥35%) · Errático (resto) · Não
previsto (recebido sem previsão).

## 10. Arquitetura

Híbrido: **motor Python validável** → tabela-fato tidy + **dashboard HTML** +
**pasta Excel** + **JSON** para o dashboard.

| Componente | Arquivo |
|---|---|
| Motor (ETL + modelo + métricas + views + exports) | `forecast_accuracy.py` |
| Payload + render do dashboard | `dashboard_build.py`, `dashboard_template.html` |
| Rotina mensal / entry point do .exe | `atualizar.py` |
| Build do executável | `atualizar.spec`, `build_exe.bat`, `BUILD_EXE.md` |
| Versão nativa Excel (M + DAX) | `EXCEL_POWERQUERY.md` |
| Guias em PDF (data-driven) | `guias/gerar_guias.py` |
| Documentação | `MANUAL.md`, `PLANO_ACAO.md` |

## 11. Entregáveis

Dashboard HTML autocontido · pasta Excel Power Pivot-ready · versão nativa Excel ·
executável · 3 guias em PDF · manual + rotina mensal · plano de ação · engine e
docs versionados em Git.

## 12. Critérios de aceite (definição de pronto)

1. Totais do modelo reconciliam com a soma bruta — **diferença zero ou
   integralmente explicada**.
2. **3 itens** conferidos manualmente batem (alto giro, erro grande, órfão).
3. Toda métrica exibida tem fórmula documentada e pergunta de negócio.
4. Órfãos, novos e descontinuados visíveis (não escondidos por join).
5. Atualização mensal <10 min (adicionar arquivo + rodar).
6. O dashboard responde: quanto erramos? p/ que lado? em quais itens? melhorando?
7. **Qualquer item** pode ser analisado individualmente; **qualquer lag** pode ser
   selecionado.
8. Cada ranking/padrão vem com hipótese de causa e ação.
9. Dashboard sem erros de console; séries CVD-safe.

## 13. Riscos e mitigações

| Risco | Mitigação |
|---|---|
| Defeitos nos arquivos de origem (colunas corrompidas/duplicadas) | tratados no parser; reconciliação falha se algo mudar |
| Interpretar timing como erro de previsão | leitura acumulada + offset +1 mês |
| Poucos meses fechados (janela curta) | limiares reestimáveis; lags com <2 meses omitidos |
| Unidades mistas (UN/G/ML/KG) | comparação por item; validar unidade na Etapa de validação |
| Dependência de Python p/ o usuário | versão nativa Excel + .exe |

## 14. Fora de escopo (agora)

Geração de forecast; camada financeira com câmbio; integração ERP online;
multiusuário/hospedagem; leitura de demanda real (faturamento/consumo) no lugar do
recebimento físico.

## 15. Roadmap / evolução

- Camada financeira (valor × câmbio) para priorização por R$.
- Estoque de segurança sugerido por classe de acurácia.
- Ficha do item como PivotTable nativa no Excel.
- Versão em inglês dos guias para a matriz.
- Reestimar limiares e recomputar o offset a cada trimestre.
