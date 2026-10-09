# Automação da Programação (Programacao_AAAA_MM.xlsm)

Este diretório é a **fonte oficial** do código VBA da planilha de Programação.
Não mantenha cópias do código dentro de células da planilha (as antigas abas
`INSTALAR_AUTOMACAO` e `ATIVAR_STATUS_AUTO` tinham cópias que divergiram do
módulo instalado e causaram o erro da lista suspensa da coluna H).

| Arquivo | Onde colar no Editor VBA (Alt+F11) |
|---|---|
| `Modulo1.bas` | Módulos > `Módulo1` (substituir todo o conteúdo) |
| `EstaPastaDeTrabalho.cls` | Microsoft Excel Objetos > `EstaPastaDeTrabalho` (substituir todo o conteúdo) |

Depois de colar: **Depurar > Compilar VBAProject** e salvar como `.xlsm`.

## Pré-requisitos na planilha

- Aba `LISTAS`: a lista de status é a Tabela `tblStatus`, coluna `Status`.
- Nome definido `ListaStatus` (Fórmulas > Gerenciador de Nomes) = `=tblStatus[Status]`.
- A lista suspensa da coluna H e as cores dos status usam **somente** esse nome.
  A cor de cada status é a cor da própria célula na aba `LISTAS`.

## Macros (Alt+F8)

- `InstalarAutomacao` – atualiza as OPs a partir das pastas de extração e
  reaplica fórmulas e a lista suspensa (também roda ao abrir o arquivo).
- `ReaplicarListaEFormatacao` – reaplica a lista suspensa e recria a formatação
  condicional (cores do Status, fim de semana e "X" do calendário) nas abas de
  processo. Rode ao incluir/alterar um status ou se as cores ficarem erradas.

## Ao alterar o código

1. Edite aqui, no repositório.
2. Atualize `VERSAO_AUTOMACAO` no topo do `Modulo1.bas`.
3. Cole na planilha. A versão aparece no título das mensagens da automação,
   o que permite confirmar qual código está instalado.
4. Abas de processo novas ou renomeadas: ajuste só a função `AbasOperacionais`.
