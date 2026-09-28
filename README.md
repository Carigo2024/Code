# IA de PCP: planejamento de produção e materiais que aprende a cada ciclo

Esta IA decide **o que produzir, o que comprar, quanto, quando e onde**. A cada
ciclo de planejamento, ela compara o que planejou com o que aconteceu de verdade
e se corrige sozinha.

## Como ela aprende

Três mecanismos de aprendizado alimentam um motor de cálculo MRP:

```
                 ┌────────────── A CADA CICLO (semana) ──────────────┐
                 │                                                   │
  demanda real ──►  REDE NEURAL ───► previsão de demanda            │
                 │  (retreinada a cada ciclo, continua de onde parou)│
                 │                                                   │
  compras        │  MEMÓRIA DE APRENDIZADO                           │
  recebidas   ───►   • lead time real de cada fornecedor             │
  produção       │   • refugo real de cada item                      │
  realizada   ───►   • eficiência real de cada máquina               │
  erros de       │   • erro da previsão ► estoque de segurança       │
  previsão    ───►                                                   │
                 │           │                                       │
                 │           ▼                                       │
                 │  MOTOR MRP (contas exatas) ─► o que / quanto /    │
                 │  + CAPACIDADE              ─► quando / onde       │
                 └───────────────────────────────────────────────────┘
                                    │
                                    ▼
                  ordens_planejadas.csv + relatorio.md
```

- **Rede neural** (`pcp_ia/rede_neural.py`): um perceptron multicamadas escrito
  do zero em numpy (ReLU, Adam, retropropagação), então dá para ler exatamente o
  que ela faz. Ela aprende tendência, sazonalidade anual e o comportamento
  recente de cada produto. A cada ciclo, é retreinada com a demanda nova sem
  perder o que já sabia (warm start).
- **Memória de aprendizado** (`pcp_ia/aprendizado.py`): corrige os parâmetros que
  o cadastro do ERP costuma ter errados, como o fornecedor que "entrega em 2
  semanas" mas na prática leva 3, a máquina que rende 82% e não 100%, e o refugo
  que ninguém cadastrou.
- **Motor MRP** (`pcp_ia/mrp.py`) e **capacidade** (`pcp_ia/capacidade.py`): fazem
  as contas clássicas de PCP (explosão da estrutura, necessidade líquida, lotes,
  lead time, alocação em máquinas com centro alternativo). Esta parte **não** é
  aprendida de propósito, porque cálculo de material precisa ser exato. A IA
  melhora os *números que entram* no cálculo.

## Resultado na empresa fictícia

`python -m pcp_ia simular` roda 52 semanas de uma fábrica de móveis cujo cadastro
está errado, como na vida real. A mesma fábrica é planejada de dois jeitos: com
o MRP tradicional (cadastro fixo e média móvel) e com a IA.

```
   Semanas |    Erro previsão    |  Nível de serviço   |  Pendências (un.)   |  Estoque (R$ mil)
           |    sem IA    com IA |    sem IA    com IA |    sem IA    com IA |    sem IA    com IA
--------------------------------------------------------------------------------------------------
   1 a 13  |      7.1%      6.2% |     43.8%     42.9% |       277       257 |      36.4      42.7
  14 a 26  |      6.9%      6.1% |     12.7%     81.1% |       355        50 |      13.8      31.7
  27 a 39  |      8.0%      5.0% |     30.8%     99.7% |       163         1 |       6.3      28.5
  40 a 52  |      8.5%      6.0% |      6.4%     97.6% |       473         6 |      16.5      31.7
```

Nas primeiras semanas as duas versões vão mal, porque a IA ainda não tem
observações. Conforme os ciclos passam, ela descobre os lead times, o refugo e a
eficiência reais, e o nível de serviço sobe para perto de 100%. O planejamento
tradicional continua errando porque confia no cadastro. O estoque da IA é maior
porque ela passa a cobrir o lead time *real*. O estoque baixo do método
tradicional vem de faltar material, não de eficiência.

## Como usar

Requer Python 3.10+ e numpy.

```bash
pip install -r requirements.txt

# 1. Ver a IA aprendendo na empresa fictícia
python -m pcp_ia simular

# 2. Rodar um ciclo de planejamento com os dados de exemplo
python -m pcp_ia planejar --dados dados_exemplo
#    → dados_exemplo/saida/relatorio.md, ordens_planejadas.csv, ...

# 3. Testes
python -m pytest
```

### Com os seus dados

1. Copie a pasta `dados_exemplo/` e troque o conteúdo das planilhas pelos dados
   da sua empresa (exportados do ERP ou preenchidos à mão). O formato de cada
   arquivo está descrito no topo de `pcp_ia/dados.py`.
2. Ajuste as suas regras em `premissas.json`.
3. Toda semana:
   - acrescente a demanda da semana em `demanda_historica.csv`;
   - acrescente o que aconteceu em `realizado_compras.csv`,
     `realizado_producao.csv` e `realizado_eficiencia.csv`;
   - atualize `estoque.csv`, `ordens_abertas.csv` e `pedidos_carteira.csv`;
   - rode `python -m pcp_ia planejar --dados sua_pasta`.

A pasta `modelo/` guarda o que a IA aprendeu (pesos da rede e previsões
passadas). **Não apague essa pasta**: é ela que faz a IA melhorar de um ciclo
para o outro.

## Premissas (`premissas.json`)

| Premissa | O que faz |
|---|---|
| `horizonte_semanas` | Quantas semanas à frente planejar |
| `semanas_congeladas` | Necessidades dentro dessa janela recebem alerta de "exige aprovação" |
| `nivel_servico` | Meta de atendimento (0,95 = 95%); define o tamanho do estoque de segurança |
| `percentil_lead_time` | Quão conservador ser com o lead time do fornecedor (90 = cobre 90% dos casos observados) |
| `usar_centro_alternativo` | Se a ordem pode ir para a máquina alternativa quando a principal lota |
| `amostras_minimas_aprendizado` | Quantas observações antes de confiar no aprendido em vez do cadastro |
| `suavizacao_aprendizado` | Peso das observações novas no refugo e na eficiência (0,3 = 30%) |
| `epocas_treino_inicial` / `epocas_treino_por_ciclo` | Quanto a rede neural treina no primeiro ciclo e nos seguintes |
| `camadas_rede` | Tamanho da rede neural (neurônios por camada oculta) |

Pedidos firmes em `pedidos_carteira.csv` sempre têm prioridade: a demanda
planejada é o maior valor entre a previsão e a carteira.

## Estrutura

```
pcp_ia/
  rede_neural.py   perceptron multicamadas (numpy puro)
  previsao.py      previsão de demanda com a rede
  aprendizado.py   lead time, refugo, eficiência, estoque de segurança
  mrp.py           motor MRP determinístico
  capacidade.py    alocação nos centros de trabalho
  planejador.py    junta tudo em um ciclo de planejamento
  simulacao.py     empresa fictícia para testar
  dados.py         leitura e gravação dos CSVs
  relatorio.py     relatório e planilhas de saída
dados_exemplo/     planilhas no formato esperado
tests/             testes automatizados
```

## Próximos passos possíveis

- Conectar direto no ERP (TOTVS, SAP, Senior…) em vez de usar CSV
- Um agente de IA (chat) em cima deste motor, para perguntar em português
  "o que preciso comprar esta semana?" ou "por que essa OP atrasou?"
- Sequenciamento fino de máquinas (ordem das OPs dentro da semana, setups)
- Premissas específicas da sua empresa (prioridade por cliente, curva ABC,
  política por fornecedor)
