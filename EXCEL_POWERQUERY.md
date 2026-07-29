# Versão Excel nativa — atualizar sem Python (Power Query + Power Pivot)

Monta um workbook que lê os arquivos crus (`Order_NN_2026.xlsx` e
`Received_itens_2026.xlsx`) **direto da pasta** e reproduz o modelo de
acuracidade. Depois de montado, a atualização mensal é só:

> **Dados → Atualizar Tudo** (nenhum Python).

Rotina mensal: jogue o novo `Order_NN_2026.xlsx` na pasta `data/`, substitua o
`Received_itens_2026.xlsx` e clique em **Atualizar Tudo**. As novas safras entram
sozinhas (a consulta lê todos os `Order_*` da pasta).

> Observação: este guia foi revisado sem um Excel para teste. O ponto mais
> sensível (defeitos dos arquivos) já vem tratado no M. Se a coluna **Alvo** não
> converter para data no seu Excel, ajuste o `Date.From` para a sua cultura
> (pt-BR) — está sinalizado no código.

---

## 0. Preparação

1. Crie uma pasta fixa, ex.: `C:\Forecast\data\`, e coloque lá os
   `Order_NN_2026.xlsx` e o `Received_itens_2026.xlsx`.
2. Abra uma pasta de trabalho nova → **Dados → Obter Dados → Consulta Nula**
   (ou *Iniciar Editor do Power Query*).
3. Em **Gerenciar Parâmetros → Novo**, crie o parâmetro **`Pasta`** (texto) =
   `C:\Forecast\data`. Todas as consultas usam esse parâmetro (troca de caminho
   num lugar só).

---

## 1. Consulta `Forecast` (grão Item × Alvo × Safra × Lag)

Cole em **Nova Consulta → Consulta Nula → Editor Avançado**. Trata os defeitos:
safra pelo nome do arquivo, coluna corrompida do Order_06 → Mai/27, coluna
Fev/27 duplicada do Order_03 descartada.

> Aceita nomes com **espaço ou underscore** (`Order 01 2026.xlsx` ou
> `Order_01_2026.xlsx`), ignora arquivos temporários `~$` e pega a data do
> **valor** do cabeçalho (não converte texto) — robusto à cultura pt-BR.

```m
let
    Arquivos = Folder.Files(Pasta),
    Orders = Table.SelectRows(Arquivos, each
        Text.StartsWith([Name], "Order")
        and Text.EndsWith(Text.Lower([Name]), ".xlsx")
        and not Text.StartsWith([Name], "~")),

    // safra = numero do mes no nome (entre "Order" e "2026"), so os digitos
    ComSafra = Table.AddColumn(Orders, "Safra", each
        try #date(2026,
            Number.FromText(Text.Select(Text.BetweenDelimiters([Name], "Order", "2026"), {"0".."9"})),
            1) otherwise null, type date),
    SoValidos = Table.SelectRows(ComSafra, each [Safra] <> null),

    Tidy = Table.AddColumn(SoValidos, "T", each
        let
            aba = Table.SelectRows(Excel.Workbook([Content], null, true),
                    each [Name]="FORECAST" and [Kind]="Sheet"){0}[Data],
            semTitulo  = Table.Skip(aba, 1),                 // tira a linha de titulo
            colNames   = Table.ColumnNames(semTitulo),
            headerVals = Record.FieldValues(semTitulo{0}),   // linha de cabecalho (datas reais)
            mapaData   = Record.FromList(headerVals, colNames),
            dados = Table.Skip(semTitulo, 1),                // linhas de item
            ren = Table.RenameColumns(dados,
                    {{colNames{0},"ItemUDB"},{colNames{2},"Desc"},{colNames{3},"UM"}}, MissingField.Ignore),
            semUPI = Table.RemoveColumns(ren, {colNames{1}}, MissingField.Ignore),
            unpiv = Table.UnpivotOtherColumns(semUPI, {"ItemUDB","Desc","UM"}, "ColMes", "Fcst"),
            comAlvo = Table.AddColumn(unpiv, "Alvo", each
                let v = Record.Field(mapaData, [ColMes]) in
                if v = null then null                                   // coluna de cabecalho vazia
                else if v is datetime or v is date then Date.From(v)
                else if Text.Contains(Text.From(v), "20267") then #date(2027,5,1)  // Order_06 corrompido
                else try Date.From(v) otherwise null, type date),
            ok = Table.SelectRows(comAlvo, each [Alvo]<>null and [Fcst]<>null and [ItemUDB]<>null)
        in
            Table.SelectColumns(ok, {"ItemUDB","Desc","UM","Alvo","Fcst"})),

    Expand = Table.ExpandTableColumn(
        Table.SelectColumns(Tidy, {"Safra","T"}), "T",
        {"ItemUDB","Desc","UM","Alvo","Fcst"}),

    Norm = Table.AddColumn(Expand, "Item", each
        let s = Text.Upper(Text.Trim(Text.From([ItemUDB]))),
            a = if Text.EndsWith(s, "-BR") then Text.Start(s, Text.Length(s)-3) else s,
            b = if Text.EndsWith(a, "-")   then Text.Start(a, Text.Length(a)-1) else a
        in b, type text),
    Lag = Table.AddColumn(Norm, "Lag", each
        (Date.Year([Alvo])*12 + Date.Month([Alvo]))
        - (Date.Year([Safra])*12 + Date.Month([Safra])), Int64.Type),

    Final = Table.SelectColumns(Lag, {"Item","Alvo","Safra","Lag","Fcst"}),
    Tipos = Table.TransformColumnTypes(Final,
        {{"Fcst", type number}, {"Alvo", type date}, {"Safra", type date}})
in
    Tipos
```

---

## 2. Consulta `Recebido` (UPI, Qtd Delivered, por Item × Alvo)

```m
let
    Arq = Table.SelectRows(Folder.Files(Pasta), each
        Text.StartsWith([Name], "Received") and not Text.StartsWith([Name], "~")),
    Conteudo = Arq{0}[Content],
    wb   = Excel.Workbook(Conteudo, null, true),
    exp  = wb{[Item="Export", Kind="Sheet"]}[Data],
    prom = Table.PromoteHeaders(exp, [PromoteAllScalars=true]),

    upi  = Table.SelectRows(prom, each Text.Trim(Text.From([Supplier])) = "UPI"),
    tipos = Table.TransformColumnTypes(upi,
        {{"Invoice Date", type date}, {"Qtd Delivered", type number}}),

    // normaliza o item na propria coluna
    norm = Table.TransformColumns(tipos, {{"Item", each
        let s = Text.Upper(Text.Trim(Text.From(_))),
            a = if Text.EndsWith(s, "-BR") then Text.Start(s, Text.Length(s)-3) else s,
            b = if Text.EndsWith(a, "-")   then Text.Start(a, Text.Length(a)-1) else a
        in b, type text}}),

    alvo = Table.AddColumn(norm, "Alvo", each Date.StartOfMonth([Invoice Date]), type date),
    sel  = Table.SelectColumns(alvo,
        {"Item", "Alvo", "Qtd Delivered", "Item Description", "Family", "Brand"}),
    filt = Table.SelectRows(sel, each [Qtd Delivered] <> null and [Alvo] <> null),

    agr  = Table.Group(filt, {"Item", "Alvo"},
        {{"Recv", each List.Sum([#"Qtd Delivered"]), type number}})
in
    agr
```

Crie também `DimItem` (dimensão do item) a partir do recebido:

```m
let
    src = Table.SelectRows(Folder.Files(Pasta), each Text.StartsWith([Name],"Received") and not Text.StartsWith([Name],"~")){0}[Content],
    exp = Table.PromoteHeaders(Excel.Workbook(src, null, true){[Item="Export",Kind="Sheet"]}[Data], [PromoteAllScalars=true]),
    upi = Table.SelectRows(exp, each Text.Trim(Text.From([Supplier]))="UPI"),
    norm = Table.TransformColumns(upi, {{"Item", each
        let s=Text.Upper(Text.Trim(Text.From(_))),
            a=if Text.EndsWith(s,"-BR") then Text.Start(s,Text.Length(s)-3) else s,
            b=if Text.EndsWith(a,"-") then Text.Start(a,Text.Length(a)-1) else a in b, type text}}),
    sel = Table.SelectColumns(norm, {"Item","Item Description","Family","Brand"}),
    dist = Table.Distinct(sel, {"Item"})
in
    dist
```

---

## 3. Calendário (`DimCalendario`)

```m
let
    ini = #date(2026,1,1),
    dias = List.Dates(ini, 24*31, #duration(1,0,0,0)),
    t = Table.FromList(dias, Splitter.SplitByNothing(), {"Date"}),
    mes = Table.AddColumn(t, "Mes", each Date.StartOfMonth([Date]), type date),
    ano = Table.AddColumn(mes, "Ano", each Date.Year([Date]), Int64.Type),
    mesano = Table.AddColumn(ano, "MesAno", each Date.ToText([Date], "yyyy-MM"), type text),
    trim = Table.AddColumn(mesano, "Trimestre", each
        Text.From(Date.Year([Date])) & "Q" & Text.From(Date.QuarterOfYear([Date])), type text)
in
    Table.Distinct(Table.SelectColumns(trim, {"Mes","Ano","MesAno","Trimestre"}), {"Mes"})
```

Em **Fechar e Carregar Para…**, marque **Somente Criar Conexão** + **Adicionar ao
Modelo de Dados** para as 4 consultas.

---

## 4. Relacionamentos (Power Pivot → Exibição de Diagrama)

- `DimItem[Item]` 1 — * `Forecast[Item]`
- `DimItem[Item]` 1 — * `Recebido[Item]`
- `DimCalendario[Mes]` 1 — * `Forecast[Alvo]`
- `DimCalendario[Mes]` 1 — * `Recebido[Alvo]`

Coloque **`Forecast[Lag]`** como Segmentação (slicer) — é ela que define a leitura
(selecione **2** para a leitura honesta; mude para ver a sensibilidade por lag).

---

## 5. Medidas DAX (Power Pivot → Nova Medida)

```DAX
Previsto := SUM ( Forecast[Fcst] )            -- respeita o slicer de Lag

Recebido := SUM ( Recebido[Recv] )

Erro := [Recebido] - [Previsto]

Bias := DIVIDE ( [Recebido] - [Previsto], [Previsto] )

-- (c) acumulada: soma F e A por item na seleção e só então o |erro| -> VOLUME
WAPE_acum :=
DIVIDE (
    SUMX ( VALUES ( DimItem[Item] ), ABS ( [Recebido] - [Previsto] ) ),
    [Recebido]
)

Acuracia := 1 - [WAPE_acum]

-- (a) mês a mês: |erro| no grão item × mês
WAPE_mes :=
DIVIDE (
    SUMX (
        CROSSJOIN ( VALUES ( DimItem[Item] ), VALUES ( DimCalendario[Mes] ) ),
        ABS ( [Recebido] - [Previsto] )
    ),
    [Recebido]
)

-- (b) trimestral
WAPE_trim :=
DIVIDE (
    SUMX (
        CROSSJOIN ( VALUES ( DimItem[Item] ), VALUES ( DimCalendario[Trimestre] ) ),
        ABS ( [Recebido] - [Previsto] )
    ),
    [Recebido]
)

-- consolidado (última safra por item×alvo) — leitura gerencial, ignora o slicer de Lag
Previsto_Consolidado :=
SUMX (
    SUMMARIZE ( Recebido, DimItem[Item], DimCalendario[Mes] ),
    VAR maxS =
        CALCULATE ( MAX ( Forecast[Safra] ), ALL ( Forecast[Lag] ) )
    RETURN
        CALCULATE ( SUM ( Forecast[Fcst] ), Forecast[Safra] = maxS, ALL ( Forecast[Lag] ) )
)

-- classificação de padrão de erro (mesmos limiares do motor)
Classe :=
VAR b = [Bias]
VAR wm = [WAPE_mes]
VAR wc = [WAPE_acum]
RETURN
SWITCH (
    TRUE (),
    ISBLANK ( [Recebido] ), "",
    wc <= 0.15 && wm <= 0.35, "Sob controle",
    b <= -0.20, "Vies superprevisao",
    b >= 0.20,  "Vies subprevisao",
    ABS ( b ) < 0.20 && wc <= 0.25 && wm >= 0.35, "Erro de timing",
    "Erratico"
)
```

Formate `Bias`, `WAPE_*` e `Acuracia` como **Percentual**.

---

## 6. Montar as telas (Tabela Dinâmica + Segmentações)

Insira **Tabela Dinâmica → Usar o Modelo de Dados**. Segmentações:
**`Forecast[Lag]`** (fixe em 2), **`DimItem[Family]`**, **`DimCalendario[MesAno]`**
(selecione a janela honesta **2026-03…2026-06** para a leitura honesta).

- **Visão geral:** medidas `[Previsto]`, `[Recebido]`, `[Bias]`, `[WAPE_acum]`,
  `[Acuracia]` sem linhas (cartões), + gráfico de colunas por `MesAno`.
- **Tabela mestre:** `DimItem[Item]` (e Descrição) nas linhas; medidas
  `[Recebido]`, `[Previsto]`, `[Bias]`, `[WAPE_acum]`, `[Acuracia]`, `[Classe]`
  nos valores. Ordene por `[Recebido]` e filtre Top 15 para os rankings.
- **Por família:** `DimItem[Family]` nas linhas + as mesmas medidas.
- **Sensibilidade por lag:** ponha `Forecast[Lag]` nas linhas e `[Bias]`,
  `[WAPE_acum]` nos valores (tire o Lag da segmentação nessa tabela).

> A janela honesta (só meses com lag-2) é garantida filtrando `MesAno` em
> 2026-03…2026-06. Para o consolidado, use a medida `[Previsto_Consolidado]`
> no lugar de `[Previsto]` e libere a segmentação de Lag.

---

## 7. Atualização mensal (sem Python)

1. Copie o novo `Order_NN_2026.xlsx` para `C:\Forecast\data\`.
2. Substitua o `Received_itens_2026.xlsx`.
3. **Dados → Atualizar Tudo.**
4. Estenda a janela honesta na segmentação `MesAno` conforme novos meses fecham
   (o lag-2 de um mês T fica disponível quando a safra T−2 já existe).

Pronto — nenhuma dependência de Python. O motor Python continua servindo para
gerar o **dashboard HTML** quando você quiser o visual pronto para gestores, mas
não é necessário para a rotina de atualização.
