# Atualizar Datas da Programação (VBA)

Macro para preencher automaticamente, na planilha de programação mensal, as datas
de **início** e **finalização** de cada O.P. a partir das extrações diárias de
`OP_ABERTAS` e `OP_FECHADAS`.

> **Nota:** nenhum código VBA foi anexado a este repositório. Esta macro foi escrita
> do zero a partir da descrição da tarefa, já tratando as causas mais comuns do erro
> *"Erro de definição de aplicativo ou de definição de objeto"* (erro 1004 do VBA).

## O que a macro faz

Para cada O.P. digitada na coluna **C** da planilha de programação:

1. **Coluna J (Data de Início)** ← data da coluna **Emissão** do *último* arquivo salvo
   na pasta `OP_ABERTAS`.
2. **Coluna K (Data de Finalização)** ← data da coluna **Encerramento** do *último*
   arquivo salvo na pasta `OP_FECHADAS`.
3. **Conferência:** compara a Data de Início (J) com a **Emissão** registrada no arquivo
   de `OP_FECHADAS` para a mesma O.P. Se as datas divergirem, a célula J é destacada em
   vermelho claro e uma observação é escrita na coluna **L**.

Ao final, um resumo é exibido (quantas datas foram preenchidas, O.P.s não encontradas e
divergências).

## Como instalar

1. Abra a planilha de programação.
2. `Alt + F11` para abrir o editor VBA.
3. Menu **Arquivo → Importar Arquivo…** e selecione `AtualizarDatasProgramacao.bas`
   (ou crie um novo Módulo e cole o conteúdo).
4. Salve a planilha como **`.xlsm`** (Pasta de Trabalho Habilitada para Macro).
5. Rode a macro por `Alt + F8` → `AtualizarDatasProgramacao`.

## Ajustes antes de rodar (topo do arquivo `.bas`)

| Constante | Para que serve |
|-----------|----------------|
| `PASTA_ABERTAS` / `PASTA_FECHADAS` | Caminhos das pastas de extração |
| `ABA_PROGRAMACAO` | Nome da aba onde estão as O.P.s |
| `PRIMEIRA_LINHA_PROG` | Primeira linha de dados (padrão 2, assumindo cabeçalho na linha 1) |
| `COL_OP_PROG` / `COL_INICIO_PROG` / `COL_FIM_PROG` | Colunas C / J / K |
| `COL_OBS_PROG` | Coluna para observações da conferência (padrão L) |
| `EXTENSAO_EXTRACAO` | Extensão dos arquivos de extração (`*.xls*`, `*.csv`, etc.) |

> As colunas **OP / Emissão / Encerramento** dentro dos arquivos de extração são
> localizadas **pelo nome do cabeçalho**, não pela letra da coluna. Assim a macro
> continua funcionando mesmo que essas colunas mudem de posição.

## Por que evita o erro 1004

O erro *"Erro de definição de aplicativo ou de definição de objeto"* costuma vir de:

- Referenciar uma aba/intervalo que não existe → **aqui há verificação da aba e das pastas**.
- Usar `.Find` e acessar `.Row` de um resultado `Nothing` → **tratado com checagem de `Nothing`**.
- Abrir/ler um arquivo que não existe → **confere se há arquivo antes de abrir**.
- Fixar letra de coluna que mudou de lugar → **detecção de coluna por cabeçalho**.
- Deixar o Excel com `ScreenUpdating`/`Calculation` alterados após um erro →
  **rotina de limpeza restaura tudo e fecha os arquivos, mesmo em caso de erro**.

## Recomendação de nome do arquivo

Como a planilha é renomeada todo mês, recomendo o padrão:

```
Programacao_AAAA_MM.xlsm      ->  ex.: Programacao_2026_08.xlsm
```

- **Sem acentos** (`Programacao`, não `Programação`) — acentos causam problemas em
  caminhos e macros.
- **Ano antes do mês** e **mês como número** (`2026_08`) — ordena cronologicamente e
  não depende do idioma.
- **Underscore no lugar de espaço** — mais seguro que `Programacao 2026_08`.

Vantagem: a própria macro pode montar o nome do mês atual sozinha, por exemplo
`"Programacao_" & Format(Date, "yyyy_mm") & ".xlsm"`.
