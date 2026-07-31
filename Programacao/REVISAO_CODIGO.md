# Revisão do código `InstalarAutomacao`

Análise do código VBA enviado, com foco no erro da imagem:
**"Não foi possível concluir a instalação: Erro de definição de aplicativo ou de definição de objeto"** (erro **1004** do VBA).

- Código original: `InstalarAutomacao_ORIGINAL.bas` (referência)
- Código corrigido: `InstalarAutomacao.bas`

---

## Causa mais provável do erro 1004 (🔴 crítico)

O erro 1004 — *"Erro de definição de aplicativo ou de definição de objeto"* — quase
sempre aparece ao **escrever em uma planilha PROTEGIDA**. O seu código escreve/edita
em várias abas durante a instalação, sem desprotegê-las antes:

| Onde | O que faz | Gera 1004 se a aba estiver protegida? |
|------|-----------|:--:|
| `ConfigurarListaStatus` | `wsLista.Cells(...).Value = ...` | ✅ |
| `ConfigurarListaStatus` | `rng.Validation.Delete` / `.Add` | ✅ |
| `AtualizarStatusFinalizado` | `ws.Cells(linha, 8).Value = "FINALIZADO"` | ✅ |
| `GravarBase` | `ws.Range("A:G").ClearContents` e escrita da base | ✅ |

Como a mensagem começa com *"Não foi possível concluir a instalação"* (que é o handler
de `InstalarAutomacao`), e a instalação chama exatamente essas três rotinas, **abas de
produção protegidas são a explicação mais provável**.

**Correção aplicada:** helpers `Desproteger` / `Reproteger`. Cada aba é desprotegida
antes da escrita e reprotegida depois, usando `UserInterfaceOnly:=True` (o usuário
continua sem conseguir editar manualmente, mas o VBA escreve sem erro). Se as abas
tiverem senha, informe-a na constante `SENHA_PROTECAO` no topo do módulo.

---

## Outros problemas encontrados

### 🔴 2. Aba inexistente derruba a rotina (erro 9)
`Set ws = ThisWorkbook.Worksheets(CStr(nomes(i)))` gera **erro 9 "Subscript out of range"**
se qualquer uma das 11 abas (`Formulacao`, `Enchimento-Forma`, `Cord`, `Go`,
`Indumak-Minipack`, `Manual`, `Montagem Valo`, `Nacionalizacao`, `Universal`,
`Embalagem`, `Apoio`) não existir ou estiver com o nome ligeiramente diferente.
O mesmo vale para `INSTALAR_AUTOMACAO`, `LISTAS` e `_BASE_OP`.
**Correção:** função `PlanilhaExiste` — abas ausentes são ignoradas (nas listas) ou
geram mensagem clara (`LISTAS`, `_BASE_OP`).

### 🟠 3. Mensagens de erro não diziam ONDE falhou
O handler mostrava só `Err.Description` ("Erro de definição de aplicativo…"), sem indicar
a etapa nem o `Err.Number`. Por isso ficou difícil localizar a falha.
**Correção:** variável `gEtapa` atualizada a cada passo; as mensagens agora mostram
**etapa + Err.Number + descrição**.

### 🟠 4. Valores de erro na coluna K podiam gerar erro 13
Em `AtualizarStatusFinalizado`, se a célula K contém um erro de fórmula (`#N/A`, `#VALOR!`),
`CStr(valorData)` gera **erro 13 (tipo incompatível)**.
**Correção:** `If Not IsError(raw) Then ...` antes de qualquer conversão.

### 🟡 5. Pastas de extração sem verificação
Se a pasta OneDrive não estiver sincronizada, o código só dizia "nenhum arquivo válido".
**Correção:** `PastaExiste` com mensagem específica de "pasta não encontrada".

### 🟡 6. `NormalizarOP` trunca O.P. com mais de 6 dígitos — **CONFIRME**
```vba
Else
    NormalizarOP = "0" & Left$(digitos, 5)   ' pega só os 5 PRIMEIROS dígitos
End If
```
Uma O.P. como `1234567` vira `012345`. Se as suas O.P.s podem ter 7+ dígitos reais,
isso causa correspondência errada entre planilha e extrações. Mantive o comportamento
original (para não desalinhar com o que já está gravado), mas **deixei um aviso no
código**. Se o certo for manter todos os dígitos, me avise que ajusto.

### 🟡 7. `Auto_Open` mostrava erro a cada abertura
Se as pastas estiverem indisponíveis ao abrir o arquivo, aparecia uma caixa de erro.
**Correção:** `Auto_Open` passou a engolir a falha silenciosamente (a atualização
pode ser refeita manualmente).

---

## O que **não** era bug (conferido)

- **Índices dos `Array` dos dicionários** estão coerentes entre `LerPasta` e
  `GravarBase` (abertas = `emissao, dataArquivo, nome`; fechadas =
  `emissao, encerramento, dataArquivo, nome`). A comparação de "mais recente"
  usa o índice certo em cada caso.
- **`ObterDataArquivo`**: o fluxo de `On Error GoTo UsarModificacao` e o fall-through
  para o rótulo funcionam corretamente.
- **`AutomationSecurity = 3`** (ForceDisable) é o valor adequado para ler os arquivos
  de extração com macros desativadas.

---

## Recomendação de nome do arquivo (sua pergunta)

Sim, vale padronizar. Recomendo **`Programacao_AAAA_MM.xlsm`** (ex.: `Programacao_2026_08.xlsm`):
sem acentos, ano antes do mês, mês como número e underscore no lugar de espaço.
Assim o nome ordena cronologicamente e pode ser montado por código com
`"Programacao_" & Format(Date, "yyyy_mm")`.
