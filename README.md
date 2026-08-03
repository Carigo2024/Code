# Análise de Reabastecimento — MRP Protheus

Ferramenta que, a partir do **relatório do MRP** (Protheus, aba `Resultados`) e do
**arquivo de estoque** (aba `Listagem do Browse`) de um mês, decide **item a item**
se é preciso **comprar/produzir** agora e **qual volume** — respeitando o **momento**
(mês a mês, nunca médias) em que cada entrada, saída e necessidade ocorre, e o
**lead time de cada item**.

Substitui a antiga planilha VBA (`MRP_Analise_Auto.xlsm`), cujos erros confirmados
**não** se repetem aqui (ver "O que mudou" abaixo).

## Como rodar (refresh mensal)

```bash
python3 executar.py <relatorio_mrp.xlsx> <estoque.xlsx> [saida.xlsx]
```

Lê os dois arquivos do mês, recalcula todas as decisões e exporta o resultado.
Os nomes dos arquivos variam a cada mês — por isso são argumentos. Se a estrutura
de qualquer arquivo não bater com o esperado, o programa **aborta com mensagem
clara** em vez de gerar dados errados.

## Estrutura

| Arquivo | Responsabilidade |
|---|---|
| `mrp/config.py` | Parâmetros de negócio (armazéns, mapa Tipo→ação, SKUs de revisão, limiares) |
| `mrp/leitor_mrp.py` | Parser resiliente do relatório (blocos por "Produto", séries por rótulo) |
| `mrp/leitor_estoque.py` | Agregação do estoque (Σ Qtde−Empenho, armazéns 01/05/22) |
| `mrp/motor.py` | Motor de decisão determinístico + justificativa |
| `mrp/exportador.py` | Exportação para Excel (**layout preliminar**, ver abaixo) |
| `executar.py` | CLI de refresh mensal |
| `tests/validar.py` | Diagnóstico do pipeline nos arquivos de referência |

## Lógica de decisão (determinística e auditável)

1. **Estoque-base** = disponível físico hoje: `Σ (Quantidade − Empenho)` nos
   armazéns **01, 05, 22** (Empenho é considerado já consumido).
2. **Saldo projetado mês a mês** a partir desse estoque-base, somando as
   `+ Entradas` programadas e subtraindo as saídas
   (`- Saídas` + `- Saída Estrutura` + `- Transf. Saída` − `+ Trasnf. Entrada`),
   cada uma no **seu** mês.
3. **Ruptura** = primeiro mês em que o saldo cai abaixo do **Estoque Segurança**
   (piso vindo do próprio relatório — calculado em outro estágio a partir do erro
   de previsão; **não** recalculado aqui).
4. **Volume** = déficit para recompor o Estoque Segurança no mês mais crítico,
   arredondado a **Lote Econômico** (múltiplo) e **Lote Mínimo** (piso).
5. **Lead time** = usado **só como urgência** (não desloca a curva): urgente
   quando a ruptura ocorre antes do que o lead time do item conseguiria repor.
6. **Cruzamento com o Protheus**: compara com a `Necessidade` nativa; divergências
   relevantes (inclusive "eu não peço, mas o Protheus aponta") são **destacadas**.
7. **Justificativa textual** por item, com os valores intermediários.
8. **Exceções**: `5973-BR` e `9121-P2` (importados de contrato anual / embarque
   bimestral) recebem recomendação, mas são marcados **REVISÃO MANUAL**.

Cobertura em dias = `(disponível − Estoque Segurança) / consumo diário`
(subtrai o estoque de segurança — corrige o bug do VBA que somava).

## O que mudou em relação ao VBA antigo

| Bug do VBA | Aqui |
|---|---|
| Tudo pela **média** dos 12 meses | Saldo **mês a mês**, respeitando o momento |
| Recebimentos somados num total único vs. necessidade total | Cada entrada/saída no **seu** mês, dentro da projeção |
| Cobertura **somava** o estoque de segurança | Cobertura **subtrai** o estoque de segurança |
| `Produzir` só se `Tipo=PA`, resto Comprar (sem validar) | Mapa `Tipo→ação` confirmado com o negócio |
| Estoque de uma célula divergente | Estoque físico (Empenho consumido, arm. 01/05/22), com **alerta** onde diverge do MRP |
| Sem cruzamento com o Protheus | Divergências com a `Necessidade` nativa **destacadas** |
| Importados tratados como item comum | 2 SKUs de contrato marcados para **revisão manual** |

## Pendências (a fechar com o usuário)

- **Formato de exportação**: o layout em `mrp/exportador.py` é **preliminar**.
  O formato definitivo deve seguir o **arquivo-modelo** ainda não fornecido.
- **IA**: o cálculo é 100% determinístico. Uso de IA (só na camada de
  justificativa e/ou apoio ao parser) está em avaliação — ver discussão no
  histórico do projeto.
- **Validação final**: comparar com o julgamento manual do usuário numa amostra
  (incluindo os 2 SKUs de revisão manual) antes do uso mensal recorrente.

Os arquivos de dados reais da empresa **não** são versionados (ver `.gitignore`).
