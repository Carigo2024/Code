# Acuracidade de forecast de importação por LAG

Compara o forecast de disponibilidade para embarque (arquivos `Order_0N`) com o
recebimento físico (arquivo `Received`), medindo o erro de previsão em função da
**defasagem (lag)** entre a safra do forecast e o mês-alvo.

Motor: [`forecast_lag.py`](forecast_lag.py).

```bash
python forecast_lag.py                       # roda tudo com defaults (data/)
python forecast_lag.py --lag 2               # força o lag fixo
python forecast_lag.py --csv grao.csv        # exporta o grão tidy
python forecast_lag.py --no-plot             # só o relatório em texto
```

Os `.xlsx` de origem **não** são versionados (dados proprietários). Coloque-os
em `data/` com nomes `Order_01_2026.xlsx … Order_04_2026.xlsx` e
`Received_itens_2026.xlsx`.

---

## 1. Grão do modelo

    Item  ×  Mês-alvo  ×  Safra do forecast (mês em que a previsão foi gerada)

Desse grão deriva o **Lag** = `mês-alvo − safra` (em meses):

| Lag | Significado |
|----:|-------------|
| 0 | previsão feita no próprio mês-alvo |
| 1 | feita 1 mês antes |
| … | … até 11 |

O forecast dá, por item, a quantidade *disponível para embarque na UPI* em cada
mês futuro. O recebido é a **data de entrada física** (Invoice Date). São o mesmo
evento conceitual (chegada à UPI) — é isso que torna a comparação legítima; e é
também por isso que eles **não** batem mês a mês (timing de chegada).

### Premissa de datação da safra (explícita, não escolhida em silêncio)

Os metadados dos arquivos foram sobrescritos por re-gravações, então a safra é
datada pelo número do arquivo: **`Order_0N` → mês N de 2026** (Jan, Fev, Mar, Abr).
Isso preserva o grão "uma safra por mês". Duas observações sobre os dados:

- `Order_03` começa em **Abr/26** (não cobre Março) e traz uma coluna **Fev/27
  duplicada** — a duplicata é descartada no parser (mantém a 1ª ocorrência).
- Datar a safra pela *primeira coluna* de cada arquivo colapsaria `Order_03` e
  `Order_04` na mesma safra (ambos começam em Abr), o que o grão proíbe. Por isso
  a datação por número do arquivo é a mais defensável. Para alterar, edite
  `SAFRAS_PADRAO` em `forecast_lag.py`.

---

## 2. Os dois modos de leitura

### Lag fixo (honesto) — **N = 2**

Compara sempre a previsão feita com **N meses de antecedência**. O N sai do
**lead time de importação real**, medido no próprio arquivo de recebidos como
`Invoice Date − Purchase Order Date`:

> mediana ≈ **61–66 dias (~2 meses)**; 98 % das entregas entre 1 e 3 meses.

Ou seja: para um item estar disponível no mês *T*, a PO precisou ser colocada
~2 meses antes. **A previsão que realmente travou a compra do mês *T* é a de lag 2.**
É a medida honesta — mede o forecast no ponto em que a decisão era irreversível.

### Consolidado (gerencial, otimista)

Compara com a **previsão mais recente** que cobre cada mês-alvo. Boa para leitura
de gestão, mas **otimista**: usa informação que só existia *depois* de a compra já
ter sido fechada. Não mede o forecast que dirigiu a decisão — mede "quão boa é a
nossa crença mais atual". Serve para reportar, não para responsabilizar o processo.

---

## 3. Tratamento crítico da defasagem física

O recebido é entrada física, **não demanda**. O mesmo erro de volume aparece
distorcido conforme o timing de chegada. Por isso o erro é medido em três leituras:

| Leitura | O que faz | Uso |
|---|---|---|
| **(a) mês a mês** | erro no grão item × mês | diagnóstico — **superestima** o erro (pune timing) |
| **(b) buckets trimestrais** | agrega item × trimestre antes do erro | meio-termo operacional (absorve deslize de ±1 mês) |
| **(c) acumulada** | total do item na janela inteira | **isola o erro de VOLUME** (timing some) |

Métricas: **Bias de volume** = `ΣRecebido / ΣPrevisto − 1` (sinal = viés) e
**WAPE** = `Σ|Previsto − Recebido| / ΣRecebido`.

---

## 4. Resultados (188 itens comparáveis; alvo Mar–Jun/2026)

### Leitura honesta — Lag 2

| Métrica | Valor |
|---|---|
| Bias de volume | **+18,6 %** (chegou mais do que a previsão travada) |
| WAPE (a) mês a mês | 54,6 % |
| WAPE (b) trimestral | 28,1 % |
| WAPE (c) acumulada | **21,8 %** |

O detalhe mensal mostra o porquê da distorção:

| Mês | Previsto | Recebido | erro mês | erro **acumulado** |
|---|---:|---:|---:|---:|
| 2026-03 | 1.303.137 | 1.908.304 | +46,4 % | +46,4 % |
| 2026-04 | 1.990.478 | 1.399.712 | −29,7 % | +0,4 % |
| 2026-05 | 935.341 | 2.177.030 | +132,8 % | +29,7 % |
| 2026-06 | 1.663.348 | 1.503.575 | −9,6 % | +18,6 % |

O erro mês a mês oscila de **−30 % a +133 %**; acumulado **converge para +18,6 %**.
Mais da metade do "erro" mensal (54,6 % → 21,8 %) é **timing de chegada, não volume**.

### Consolidado (mesma janela, para comparação justa)

| Métrica | Consolidado | Lag 2 |
|---|---:|---:|
| Bias de volume | −13,5 % | +18,6 % |
| WAPE acumulada | 19,4 % | 21,8 % |

O consolidado parece um pouco melhor no volume (19,4 % vs 21,8 %) e **inverte o
sinal do viés**: as safras recentes (corte de Abr) inflaram os meses próximos,
puxando o previsto acima do recebido. É exatamente o otimismo esperado — não use
para cobrar acurácia.

### Sensibilidade por lag

| Lag | Bias | WAPE mês | WAPE acum |
|----:|-----:|---------:|----------:|
| 0 | −31,7 % | 119 % | 58 % |
| 1 | +16,2 % | 63 % | 18,6 % |
| **2** | **+18,6 %** | **55 %** | **21,8 %** |
| 3 | +6,6 % | 73 % | 48 % |
| 4 | −3,4 % | 84 % | 49 % |

Lag 1–2 são o ponto ótimo. Lag 1 é marginalmente melhor, mas com lead time de
~2 meses a compra já está colocada no lag 1 → **lag 2 é a última previsão
acionável**, e por isso a honesta.

---

## 5. Recomendação

**Para julgar VOLUME (a pergunta "o forecast acerta quanto entra?"), use a leitura
(c) acumulada.** Ela neutraliza o timing de entrada física — que não é culpa do
forecast — e isola o erro real de volume. Nos dados: **~+19 % de sub-previsão** no
ponto de decisão (lag 2). A leitura **(b) trimestral** é o melhor meio-termo para
acompanhamento operacional. A **(a) mês a mês** deve ficar só como diagnóstico:
com recebido = data de entrada, ela superestima o erro (aqui, 55 % vs 22 % reais).

**Modo de referência:** o **lag fixo N = 2** para responsabilizar o processo (medida
honesta); o **consolidado** apenas como leitura gerencial, sempre rotulado como
otimista.

---

## 6. Cobertura e ressalvas

- **265** itens no forecast, **329** no recebido, **188** comparáveis (interseção).
  Códigos casados após normalizar sufixo `-BR` e traço final.
- Itens previstos e nunca recebidos, ou recebidos sem forecast, ficam fora do
  score (entram como zero do outro lado só dentro da janela avaliada).
- Julho/2026 está parcial (dados até 27/07) — por isso a janela honesta encerra em
  **Junho**, onde o lag 2 tem cobertura completa.
- Quantidades comparadas na unidade de cada item (UM do forecast = `UN`); eventual
  divergência de unidade por item não é tratada.
