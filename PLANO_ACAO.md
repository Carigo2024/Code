# Plano de ação — reduzir o erro do forecast (UPI)

Baseado nos padrões medidos (187 itens comparáveis, leitura honesta lag-2,
janela Mar–Jun/2026). Números do motor; reestimar a cada ciclo.

---

## Diagnóstico-síntese

1. **O viés é de SUBprevisão, não de superprevisão.** Recebido a mais que o
   previsto = **1.306.888 un** (85 itens); previsto a mais = **207.638 un** (56
   itens). Bias líquido **+18,7%**. Prevemos sistematicamente **abaixo** do que entra.
2. **Dói na ponta A.** Dos 23 itens ABC A (80% do volume), **13 são viés de
   subprevisão** e só 1 de superprevisão. Ou seja, o erro está concentrado nos
   itens de maior giro → **risco de ruptura e frete aéreo emergencial** onde mais custa.
3. **60% do erro mensal é timing, não volume.** WAPE mês 54,6% → acumulada 21,7%.
   50 itens são "erro de timing" (volume bate, meses não), por causa da defasagem
   física de **+1 mês**.
4. **A previsão só é confiável até ~lag 1–2.** WAPE acumulada por lag: lag 1 =
   15,8%, lag 2 = 21,7%, e degrada forte a partir do lag 3 (38,8% → 40,7% → 65,7%).
5. **Churn altíssimo em spares/equipamento.** Os itens de maior instabilidade
   (VALO batteries, LED boards, thermal grease, CRMs esporádicos) têm churn
   800–1000% e são ABC C, quase todos "não previsto".

---

## Ações priorizadas

### 1. Corrigir o viés de subprevisão dos itens A (maior ROI)
13 itens A subprevistos, liderados por:

| Item | Descrição | Vol | Bias | Acur |
|---|---|---:|---:|---:|
| 14200 | Combo Str Mac Clr Tip | 608.500 | +79,8% | 56% |
| 72112 | CRM #112 | 300.000 | +57,9% | 63% |
| 10398 | STEM 1.2ML Azul Pérola | 250.000 | +54,3% | 65% |
| 10842 | Barrel 1.2ML Transparente | 242.000 | +49,4% | 67% |
| 60088 | Plunger 1.2ML Cinza | 242.000 | +49,4% | 67% |
| 72081 | CRM #081 | 180.000 | +20,0% | 83% |

**Ação:** aplicar fator de correção de viés (bias-adjust) na previsão desses itens
— reponderar para cima pelo bias medido — ou revisar a premissa de consumo com o
demand planning. **Meta:** trazer o bias dos itens A para dentro de ±20% (sair da
classe "viés") em 2–3 ciclos. Enquanto não corrigido, **subir o estoque de
segurança** desses itens (risco de ruptura).

### 2. Estoque de segurança por classe de acurácia (não uniforme)
- **Sob controle (25 itens):** SS mínimo.
- **Viés (corrigível):** corrigir a previsão; SS transitório até o viés cair.
- **Errático (4) + timing residual + não-previsto:** cobrir com SS — *não adianta
  tentar prever melhor*.
- **Dimensionamento:** ancore o SS no erro de **volume** por classe (WAPE acum:
  ~15% sob controle, ~22% geral no lag-2, mais para erráticos). Como a previsão
  degrada além do lag 2–3, **o SS deve cobrir o erro no horizonte do lead time
  (~2 meses)** — não confie em previsão de lag longo.

### 3. Tratar os 50 itens "timing" sem mexer na previsão
Volume bate, meses não. Aplicar o **offset +1 mês** no planejamento de chegada e
planejar por **bucket trimestral**. Não gerar ordens extras nem alarmes por deslize
de 1 mês — seria corrigir o que não está quebrado.

### 4. Reduzir churn: itens intermitentes fora do forecast rolante
Os campeões de churn (equipamento/spares ABC C, churn 800–1000%) desestabilizam o
sinal enviado à matriz. **Ação:** tirar do forecast mensal e passar para
**reposição por ponto de pedido / order-to-need**. Forecast estável melhora a
credibilidade e o atendimento da matriz.

### 5. Fechar o gap dos 26 "não previsto"
Itens recebidos com previsão zero na janela. Revisar por que não entram no forecast
(novos, spares, ajuste manual) e **incluí-los** onde fizer sentido.

### 6. Pauta de governança com a matriz (UPI)
- Adotar a **leitura honesta (lag-2 acumulada)** como métrica oficial de acurácia;
  tratar a **defasagem de +1 mês como característica do fluxo**, não como erro.
- Discutir o **viés sistemático de +18,7%**: a matriz entrega mais que o pedido, ou
  o pedido subestima a necessidade? Alinhar a premissa.
- Definir **horizonte de congelamento (freeze) ~2 meses**, coerente com o lead time:
  a previsão lag-2 é a que trava a compra; revisões depois disso não mudam o buy.
- Acompanhamento mensal pelo dashboard, **foco nos itens A**.

---

## Metas quantificadas (2–3 ciclos)

| Indicador | Hoje | Meta |
|---|---:|---:|
| Bias dos itens A | ~+50% (média dos viesados) | dentro de ±20% |
| WAPE acumulada geral (lag-2) | 21,7% | < 15% (mais itens "sob controle") |
| Itens com \|TS\|>4 ou churn>100% | vários | migrados p/ reposição |
| Cobertura de forecast (não-previsto) | 26 itens | endereçados |
