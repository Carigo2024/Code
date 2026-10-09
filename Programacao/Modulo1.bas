Option Explicit

'Fonte oficial deste código: repositório Git (Programacao/Modulo1.bas).
'Não guarde cópias dele em células da planilha.
Public Const VERSAO_AUTOMACAO As String = "2026-10-09"

Private Const PASTA_ABERTAS As String = "C:\Users\CEmerson\Ultradent Products Inc\PCP e Logística - Hard work - Documentos\CONTROL BOARD\EXTRACOES\OP_ABERTAS"
Private Const PASTA_FECHADAS As String = "C:\Users\CEmerson\Ultradent Products Inc\PCP e Logística - Hard work - Documentos\CONTROL BOARD\EXTRACOES\OP_FECHADAS"
Private Const ABA_BASE As String = "_BASE_OP"

'Nome definido (Fórmulas > Gerenciador de Nomes) = tblStatus[Status] na aba LISTAS.
'A lista suspensa e as cores do Status usam somente este nome.
Private Const NOME_LISTA_STATUS As String = "ListaStatus"
Private Const ULTIMA_LINHA_STATUS As Long = 1000

Public Function AbasOperacionais() As Variant
    AbasOperacionais = Array( _
        "Formulacao", "Enchimento-Forma", "Cord", "Go", _
        "Indumak-Minipack", "Manual", "Montagem Valo", _
        "Nacionalizacao", "Universal", "Embalagem", "Apoio")
End Function

Public Sub Auto_Open()
    AtualizarProgramacao False
End Sub

Public Sub InstalarAutomacao()
    AtualizarProgramacao True
End Sub

Public Sub AtualizarProgramacao(Optional ByVal ExibirMensagem As Boolean = True)
    Dim abertas As Object
    Dim fechadas As Object
    Dim anterior As Object
    Dim qtdArquivosAbertos As Long
    Dim qtdArquivosFechados As Long
    Dim eventosAnteriores As Boolean
    Dim telaAnterior As Boolean
    Dim alertasAnteriores As Boolean
    Dim linksAnteriores As Boolean
    Dim segurancaAnterior As Long
    Dim statusAnterior As Variant
    Dim numeroErro As Long
    Dim descricaoErro As String

    On Error GoTo TrataErro

    eventosAnteriores = Application.EnableEvents
    telaAnterior = Application.ScreenUpdating
    alertasAnteriores = Application.DisplayAlerts
    linksAnteriores = Application.AskToUpdateLinks
    segurancaAnterior = Application.AutomationSecurity
    statusAnterior = Application.StatusBar

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.AutomationSecurity = 3
    Application.StatusBar = "Atualizando ordens de produção..."
    Application.Calculation = xlCalculationManual

    If Not PastaExiste(PASTA_ABERTAS) Then
        Err.Raise vbObjectError + 101, , "Pasta de OP abertas não encontrada: " & PASTA_ABERTAS
    End If

    If Not PastaExiste(PASTA_FECHADAS) Then
        Err.Raise vbObjectError + 102, , "Pasta de OP fechadas não encontrada: " & PASTA_FECHADAS
    End If

    Set abertas = CreateObject("Scripting.Dictionary")
    Set fechadas = CreateObject("Scripting.Dictionary")
    Set anterior = CarregarBaseAnterior()

    abertas.CompareMode = vbTextCompare
    fechadas.CompareMode = vbTextCompare

    qtdArquivosAbertos = LerPasta(PASTA_ABERTAS, False, abertas)
    qtdArquivosFechados = LerPasta(PASTA_FECHADAS, True, fechadas)

    If qtdArquivosAbertos = 0 Then
        Err.Raise vbObjectError + 103, , "Nenhum arquivo de OP abertas pôde ser lido."
    End If

    If qtdArquivosFechados = 0 Then
        Err.Raise vbObjectError + 104, , "Nenhum arquivo de OP fechadas pôde ser lido."
    End If

    GravarBase abertas, fechadas, anterior
    AplicarFormulasAutomacao
    ConfigurarListaStatus

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    AtualizarStatusFinalizado

Finalizar:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    Application.ScreenUpdating = telaAnterior
    Application.DisplayAlerts = alertasAnteriores
    Application.AskToUpdateLinks = linksAnteriores
    Application.AutomationSecurity = segurancaAnterior
    Application.StatusBar = statusAnterior
    Application.EnableEvents = eventosAnteriores
    On Error GoTo 0

    If numeroErro <> 0 Then
        MsgBox "A atualização não foi concluída." & vbCrLf & vbCrLf & descricaoErro, _
               vbExclamation, "Automação de OP - versão " & VERSAO_AUTOMACAO
    ElseIf ExibirMensagem Then
        MsgBox "Atualização concluída." & vbCrLf & _
               abertas.Count & " OPs abertas e " & fechadas.Count & " OPs fechadas foram consolidadas.", _
               vbInformation, "Automação de OP - versão " & VERSAO_AUTOMACAO
    End If

    Exit Sub

TrataErro:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    Resume Finalizar
End Sub

Private Function PastaExiste(ByVal caminho As String) As Boolean
    Dim fso As Object

    On Error GoTo Saida
    Set fso = CreateObject("Scripting.FileSystemObject")
    PastaExiste = fso.FolderExists(caminho)

Saida:
    Set fso = Nothing
End Function

Private Function LerPasta(ByVal pasta As String, _
                          ByVal arquivoFechado As Boolean, _
                          ByRef destino As Object) As Long
    Dim nomeArquivo As String
    Dim caminhoCompleto As String

    nomeArquivo = Dir$(pasta & "\*.xls*")

    Do While Len(nomeArquivo) > 0
        If Left$(nomeArquivo, 2) <> "~$" _
           And StrComp(nomeArquivo, ThisWorkbook.Name, vbTextCompare) <> 0 Then

            caminhoCompleto = pasta & "\" & nomeArquivo

            If LerArquivo(caminhoCompleto, nomeArquivo, arquivoFechado, destino) Then
                LerPasta = LerPasta + 1
            End If
        End If

        nomeArquivo = Dir$()
    Loop
End Function

Private Function LerArquivo(ByVal caminhoCompleto As String, _
                            ByVal nomeArquivo As String, _
                            ByVal arquivoFechado As Boolean, _
                            ByRef destino As Object) As Boolean
    Dim wbFonte As Workbook
    Dim wsFonte As Worksheet
    Dim abriuAqui As Boolean
    Dim ultimaLinha As Long
    Dim linha As Long
    Dim op As String
    Dim registro As Variant
    Dim registroAnterior As Variant
    Dim carimboFonte As Double
    Dim indiceData As Long

    On Error GoTo Saida

    On Error Resume Next
    Set wbFonte = Application.Workbooks(nomeArquivo)
    On Error GoTo Saida

    If wbFonte Is Nothing Then
        Set wbFonte = Workbooks.Open(Filename:=caminhoCompleto, _
                                     UpdateLinks:=False, _
                                     ReadOnly:=True, _
                                     Notify:=False, _
                                     AddToMru:=False, _
                                     IgnoreReadOnlyRecommended:=True)
        abriuAqui = True
    End If

    Set wsFonte = wbFonte.Worksheets("Export")
    ultimaLinha = wsFonte.Cells(wsFonte.Rows.Count, 2).End(xlUp).Row
    carimboFonte = CarimboArquivo(nomeArquivo, FileDateTime(caminhoCompleto))

    If arquivoFechado Then
        indiceData = 3
    Else
        indiceData = 2
    End If

    For linha = 2 To ultimaLinha
        op = NormalizarOP(wsFonte.Cells(linha, 2).Value2)

        If Len(op) > 0 Then
            If arquivoFechado Then
                registro = Array( _
                    wsFonte.Cells(linha, 3).Value2, _
                    wsFonte.Cells(linha, 4).Value2, _
                    ValorQuantidade(wsFonte.Cells(linha, 9).Value2), _
                    carimboFonte, _
                    nomeArquivo)
            Else
                registro = Array( _
                    wsFonte.Cells(linha, 3).Value2, _
                    ValorQuantidade(wsFonte.Cells(linha, 9).Value2), _
                    carimboFonte, _
                    nomeArquivo)
            End If

            If destino.Exists(op) Then
                registroAnterior = destino(op)

                If CDbl(registro(indiceData)) > CDbl(registroAnterior(indiceData)) _
                   Or (CDbl(registro(indiceData)) = CDbl(registroAnterior(indiceData)) _
                       And StrComp(nomeArquivo, CStr(registroAnterior(indiceData + 1)), vbTextCompare) > 0) Then
                    destino(op) = registro
                End If
            Else
                destino.Add op, registro
            End If
        End If
    Next linha

    LerArquivo = True

Saida:
    On Error Resume Next
    If abriuAqui Then wbFonte.Close SaveChanges:=False
    Set wsFonte = Nothing
    Set wbFonte = Nothing
    On Error GoTo 0
End Function

Private Function CarimboArquivo(ByVal nomeArquivo As String, _
                                ByVal modificadoEm As Date) As Double
    Dim re As Object
    Dim mt As Object
    Dim ano As Long
    Dim mes As Long
    Dim dia As Long
    Dim dataEncontrada As Date

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(20[0-9]{2})[-_](0[1-9]|1[0-2])[-_]([0-2][0-9]|3[01])"

    If re.Test(nomeArquivo) Then
        Set mt = re.Execute(nomeArquivo)(0)
        ano = CLng(mt.SubMatches(0))
        mes = CLng(mt.SubMatches(1))
        dia = CLng(mt.SubMatches(2))
        dataEncontrada = DateSerial(ano, mes, dia)

        If Year(dataEncontrada) = ano _
           And Month(dataEncontrada) = mes _
           And Day(dataEncontrada) = dia Then
            CarimboArquivo = CDbl(dataEncontrada) + _
                             (CDbl(modificadoEm) - Fix(CDbl(modificadoEm)))
            Exit Function
        End If
    End If

    CarimboArquivo = CDbl(modificadoEm)
End Function

Private Function NormalizarOP(ByVal valor As Variant) As String
    Dim digitos As String
    Dim numero As Double

    If IsError(valor) Or IsEmpty(valor) Then Exit Function

    If VarType(valor) = vbString Then
        digitos = Trim$(CStr(valor))
        digitos = Replace$(digitos, " ", vbNullString)
        digitos = Replace$(digitos, ChrW(160), vbNullString)

        If Len(digitos) = 0 Or digitos Like "*[!0-9]*" Then Exit Function
    ElseIf IsNumeric(valor) And VarType(valor) <> vbBoolean Then
        numero = CDbl(valor)

        If numero < 0 Or numero <> Fix(numero) Then Exit Function
        digitos = Format$(numero, "0")
    Else
        Exit Function
    End If

    If Len(digitos) <= 6 Then
        NormalizarOP = Right$("000000" & digitos, 6)
    Else
        NormalizarOP = "0" & Left$(digitos, 5)
    End If
End Function

Private Function ValorQuantidade(ByVal valor As Variant) As Variant
    If IsError(valor) Or IsEmpty(valor) Then
        ValorQuantidade = Empty
    ElseIf Len(Trim$(CStr(valor))) = 0 Then
        ValorQuantidade = Empty
    ElseIf IsNumeric(valor) Then
        ValorQuantidade = CDbl(valor)
    Else
        ValorQuantidade = Empty
    End If
End Function

Private Function CarregarBaseAnterior() As Object
    Dim dados As Object
    Dim wsBase As Worksheet
    Dim ultimaLinha As Long
    Dim matriz As Variant
    Dim linha As Long
    Dim op As String

    Set dados = CreateObject("Scripting.Dictionary")
    dados.CompareMode = vbTextCompare

    On Error GoTo Saida

    Set wsBase = ThisWorkbook.Worksheets(ABA_BASE)
    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row

    If ultimaLinha < 2 Then GoTo Saida

    matriz = wsBase.Range("A2:I" & ultimaLinha).Value2

    For linha = 1 To UBound(matriz, 1)
        op = NormalizarOP(matriz(linha, 1))

        If Len(op) > 0 Then
            dados(op) = Array( _
                matriz(linha, 2), _
                matriz(linha, 3), _
                matriz(linha, 4), _
                matriz(linha, 5), _
                matriz(linha, 6), _
                matriz(linha, 7), _
                matriz(linha, 8), _
                matriz(linha, 9))
        End If
    Next linha

Saida:
    Set CarregarBaseAnterior = dados
End Function

Private Sub GravarBase(ByVal abertas As Object, _
                       ByVal fechadas As Object, _
                       ByVal anterior As Object)
    Dim todas As Object
    Dim wsBase As Worksheet
    Dim chave As Variant
    Dim registroAberto As Variant
    Dim registroFechado As Variant
    Dim registroAnterior As Variant
    Dim Saida() As Variant
    Dim cabecalho(1 To 1, 1 To 9) As Variant
    Dim linha As Long
    Dim ultimaLinhaAnterior As Long
    Dim emissao As Variant
    Dim encerramento As Variant
    Dim situacao As String
    Dim arquivoEmissao As String
    Dim arquivoEncerramento As String
    Dim atualizadoEm As Variant
    Dim qtdProduzida As Variant
    Dim arquivoQuantidade As String

    Set todas = CreateObject("Scripting.Dictionary")
    todas.CompareMode = vbTextCompare

    For Each chave In anterior.Keys
        todas(chave) = True
    Next chave

    For Each chave In abertas.Keys
        todas(chave) = True
    Next chave

    For Each chave In fechadas.Keys
        todas(chave) = True
    Next chave

    If todas.Count > 0 Then ReDim Saida(1 To todas.Count, 1 To 9)

    For Each chave In todas.Keys
        emissao = Empty
        encerramento = Empty
        situacao = vbNullString
        arquivoEmissao = vbNullString
        arquivoEncerramento = vbNullString
        atualizadoEm = Empty
        qtdProduzida = Empty
        arquivoQuantidade = vbNullString

        If anterior.Exists(chave) Then
            registroAnterior = anterior(chave)
            emissao = registroAnterior(0)
            encerramento = registroAnterior(1)
            situacao = CStr(registroAnterior(2))
            arquivoEmissao = CStr(registroAnterior(3))
            arquivoEncerramento = CStr(registroAnterior(4))
            atualizadoEm = registroAnterior(5)
            qtdProduzida = registroAnterior(6)
            arquivoQuantidade = CStr(registroAnterior(7))
        End If

        'Só considera fechamento quando a OP foi lida em OP_FECHADAS nesta atualização.
        If Not fechadas.Exists(chave) Then
            encerramento = Empty
            arquivoEncerramento = vbNullString

            If UCase$(Trim$(situacao)) = "FECHADA" Then
                situacao = vbNullString
            End If
        End If

        If abertas.Exists(chave) Then
            registroAberto = abertas(chave)

            If TemValor(registroAberto(0)) Then emissao = registroAberto(0)
            situacao = "ABERTA"
            arquivoEmissao = CStr(registroAberto(3))
            atualizadoEm = Now
            qtdProduzida = registroAberto(1)

            If TemValor(qtdProduzida) Then
                arquivoQuantidade = CStr(registroAberto(3))
            Else
                arquivoQuantidade = vbNullString
            End If
        End If

        If fechadas.Exists(chave) Then
            registroFechado = fechadas(chave)

            If (Not abertas.Exists(chave)) Or (Not TemValor(emissao)) Then
                emissao = registroFechado(0)
                arquivoEmissao = CStr(registroFechado(4))
            End If

            encerramento = registroFechado(1)
            situacao = "FECHADA"
            arquivoEncerramento = CStr(registroFechado(4))
            atualizadoEm = Now
            qtdProduzida = registroFechado(2)

            If TemValor(qtdProduzida) Then
                arquivoQuantidade = CStr(registroFechado(4))
            Else
                arquivoQuantidade = vbNullString
            End If
        End If

        linha = linha + 1
        Saida(linha, 1) = CStr(chave)
        Saida(linha, 2) = emissao
        Saida(linha, 3) = encerramento
        Saida(linha, 4) = situacao
        Saida(linha, 5) = arquivoEmissao
        Saida(linha, 6) = arquivoEncerramento
        Saida(linha, 7) = atualizadoEm
        Saida(linha, 8) = qtdProduzida
        Saida(linha, 9) = arquivoQuantidade
    Next chave

    Set wsBase = ThisWorkbook.Worksheets(ABA_BASE)
    ultimaLinhaAnterior = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row

    cabecalho(1, 1) = "OP"
    cabecalho(1, 2) = "Emissão"
    cabecalho(1, 3) = "Encerramento"
    cabecalho(1, 4) = "Situação"
    cabecalho(1, 5) = "Arquivo Emissão"
    cabecalho(1, 6) = "Arquivo Encerramento"
    cabecalho(1, 7) = "Atualizado em"
    cabecalho(1, 8) = "Qtd Produzida"
    cabecalho(1, 9) = "Arquivo Quantidade"

    With wsBase
        .Range("A1:I" & Application.Max(ultimaLinhaAnterior, todas.Count + 1)).ClearContents
        .Columns("A").NumberFormat = "@"
        .Range("A1:I1").Value = cabecalho

        If todas.Count > 0 Then
            .Range("A2").Resize(todas.Count, 9).Value = Saida
        End If
        .Columns("B:C").NumberFormat = "dd/mm/yyyy"
        .Columns("G").NumberFormat = "dd/mm/yyyy hh:mm"
        .Columns("H").NumberFormat = "#,##0"
    End With
End Sub

Private Function TemValor(ByVal valor As Variant) As Boolean
    If IsError(valor) Or IsEmpty(valor) Then Exit Function
    TemValor = (Len(Trim$(CStr(valor))) > 0)
End Function

Private Function UltimaLinhaDados(ByVal ws As Worksheet) As Long
    Dim ultimaOP As Long
    Dim ultimaSKU As Long

    ultimaOP = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    ultimaSKU = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    UltimaLinhaDados = Application.Max(4, ultimaOP, ultimaSKU)
End Function

Private Function LimiteComReserva(ByVal ws As Worksheet) As Long
    LimiteComReserva = Application.Max(200, UltimaLinhaDados(ws) + 50)
End Function

Public Function FormulaDescricao() As String
    FormulaDescricao = "=IF(RC[-1]="""","""",XLOOKUP(RC[-1]&"""",Tabela1[SKU]&"""",Tabela1[Produto],"""",0,1))"
End Function

Public Function FormulaQtd() As String
    FormulaQtd = "=LET(raw,RC[-4],key,IFERROR(IF(LEN(TEXT(--raw,""0""))<=6," & _
                 "RIGHT(""000000""&TEXT(--raw,""0""),6),""0""&LEFT(TEXT(--raw,""0""),5)),"""")," & _
                 "src,XLOOKUP(key,'_BASE_OP'!R2C1:R10000C1,'_BASE_OP'!R2C9:R10000C9,"""")," & _
                 "qty,XLOOKUP(key,'_BASE_OP'!R2C1:R10000C1,'_BASE_OP'!R2C8:R10000C8,"""")," & _
                 "IF(OR(key="""",src=""""),"""",qty))"
End Function

Public Function FormulaEmissao() As String
    FormulaEmissao = "=IF(RC[-7]="""","""",LET(k,IF(LEN(TEXT(RC[-7],""0""))<=6," & _
                     "RIGHT(""000000""&TEXT(RC[-7],""0""),6),""0""&LEFT(TEXT(RC[-7],""0""),5))," & _
                     "XLOOKUP(k,_BASE_OP!R2C1:R10000C1,_BASE_OP!R2C2:R10000C2,"""")))"
End Function

Public Function FormulaEncerramento() As String
    FormulaEncerramento = "=LET(raw,RC[-8],key,IFERROR(IF(LEN(TEXT(--raw,""0""))<=6," & _
                          "RIGHT(""000000""&TEXT(--raw,""0""),6),""0""&LEFT(TEXT(--raw,""0""),5)),"""")," & _
                          "fechada,COUNTIFS('_BASE_OP'!R2C1:R10000C1,key,'_BASE_OP'!R2C4:R10000C4,""FECHADA""," & _
                          "'_BASE_OP'!R2C6:R10000C6,""<>"")>0," & _
                          "dt,XLOOKUP(key,'_BASE_OP'!R2C1:R10000C1,'_BASE_OP'!R2C3:R10000C3,0)," & _
                          "IF(AND(fechada,ISNUMBER(dt),dt>0),dt,""""))"
End Function

Private Function ObterAba(ByVal nome As String) As Worksheet
    On Error Resume Next
    Set ObterAba = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0

    If ObterAba Is Nothing Then
        Err.Raise vbObjectError + 110, , _
            "A aba '" & nome & "' não foi encontrada." & vbCrLf & _
            "Se ela foi renomeada, ajuste a função AbasOperacionais no Módulo1."
    End If
End Function

Private Sub AplicarFormulasAutomacao()
    Dim nomes As Variant
    Dim nome As Variant
    Dim ws As Worksheet
    Dim limite As Long
    Dim linhaFormula As Long

    nomes = AbasOperacionais()

    For Each nome In nomes
        Set ws = ObterAba(CStr(nome))

        limite = LimiteComReserva(ws)

        With ws
            For linhaFormula = 4 To limite
                If Not .Cells(linhaFormula, 5).HasFormula _
                   And IsEmpty(.Cells(linhaFormula, 5).Value2) Then
                    .Cells(linhaFormula, 5).Formula2R1C1 = FormulaDescricao()
                End If
            Next linhaFormula

            .Range("G4:G" & limite).Formula2R1C1 = FormulaQtd()
            .Range("J4:J" & limite).Formula2R1C1 = FormulaEmissao()
            .Range("K4:K" & limite).Formula2R1C1 = FormulaEncerramento()
            .Range("G4:G" & limite).NumberFormat = "#,##0"
            .Range("J4:K" & limite).NumberFormat = "dd/mm/yyyy"
        End With
    Next nome
End Sub

Private Function IntervaloListaStatus() As Range
    Dim intervalo As Range

    On Error Resume Next
    Set intervalo = ThisWorkbook.Names(NOME_LISTA_STATUS).RefersToRange
    On Error GoTo 0

    If intervalo Is Nothing Then
        Err.Raise vbObjectError + 111, , _
            "O nome '" & NOME_LISTA_STATUS & "' não existe ou não aponta para células." & vbCrLf & _
            "Crie-o em Fórmulas > Gerenciador de Nomes, com a fórmula =tblStatus[Status]."
    End If

    If Application.WorksheetFunction.CountA(intervalo) = 0 Then
        Err.Raise vbObjectError + 112, , "A lista '" & NOME_LISTA_STATUS & "' está vazia."
    End If

    Set IntervaloListaStatus = intervalo
End Function

Private Sub ConfigurarListaStatus()
    Dim nomes As Variant
    Dim nome As Variant
    Dim ws As Worksheet
    Dim ultimaLinha As Long

    IntervaloListaStatus

    nomes = AbasOperacionais()

    For Each nome In nomes
        Set ws = ObterAba(CStr(nome))

        ultimaLinha = Application.Max(ULTIMA_LINHA_STATUS, LimiteComReserva(ws))

        With ws.Range("H4:H" & ultimaLinha).Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:="=" & NOME_LISTA_STATUS
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
            .ErrorTitle = "Status inválido"
            .ErrorMessage = "Escolha um status da lista (aba LISTAS)."
        End With
    Next nome
End Sub

'Reaplica a lista suspensa e recria a formatação condicional das abas de processo.
'Rode sempre que as cores sumirem/ficarem erradas ou ao incluir um status novo na aba LISTAS.
Public Sub ReaplicarListaEFormatacao()
    Dim abaOriginal As Object
    Dim telaAnterior As Boolean
    Dim eventosAnteriores As Boolean
    Dim avisos As String
    Dim numeroErro As Long
    Dim descricaoErro As String

    Set abaOriginal = ActiveSheet
    telaAnterior = Application.ScreenUpdating
    eventosAnteriores = Application.EnableEvents

    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ConfigurarListaStatus
    avisos = AplicarFormatacaoStatus()

Finalizar:
    On Error Resume Next
    abaOriginal.Activate
    Application.EnableEvents = eventosAnteriores
    Application.ScreenUpdating = telaAnterior
    On Error GoTo 0

    If numeroErro <> 0 Then
        MsgBox "Não foi possível reaplicar a lista e as cores." & vbCrLf & vbCrLf & descricaoErro, _
               vbExclamation, "Automação de OP - versão " & VERSAO_AUTOMACAO
    ElseIf Len(avisos) > 0 Then
        MsgBox "Lista e cores reaplicadas, com avisos:" & avisos, _
               vbExclamation, "Automação de OP - versão " & VERSAO_AUTOMACAO
    Else
        MsgBox "Lista suspensa e cores do Status reaplicadas nas abas de processo.", _
               vbInformation, "Automação de OP - versão " & VERSAO_AUTOMACAO
    End If

    Exit Sub

TrataErro:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    Resume Finalizar
End Sub

Private Function AplicarFormatacaoStatus() As String
    Dim listaStatus As Range
    Dim celula As Range
    Dim celulaFinalizado As Range
    Dim nomes As Variant
    Dim nome As Variant
    Dim ws As Worksheet
    Dim faixaStatus As Range
    Dim regra As FormatCondition
    Dim ultimaLinha As Long
    Dim formulaFimDeSemana As String
    Dim avisos As String

    Set listaStatus = IntervaloListaStatus()

    For Each celula In listaStatus.Cells
        If Len(Trim$(CStr(celula.Value2))) > 0 Then
            If celula.Interior.ColorIndex = xlColorIndexNone Then
                If Not PintarCorPadrao(celula) Then
                    avisos = avisos & vbCrLf & "- pinte a célula do status '" & _
                             celula.Value2 & "' na aba LISTAS e rode de novo."
                End If
            End If

            If UCase$(Trim$(CStr(celula.Value2))) = "FINALIZADO" Then Set celulaFinalizado = celula
        End If
    Next celula

    formulaFimDeSemana = FormulaNoIdiomaDoExcel("=AND(L3<>"""",WEEKDAY(L3,2)>=6)")

    nomes = AbasOperacionais()

    For Each nome In nomes
        Set ws = ObterAba(CStr(nome))

        ultimaLinha = Application.Max(ULTIMA_LINHA_STATUS, LimiteComReserva(ws))

        'Regras com referência relativa são criadas a partir da célula ativa.
        ws.Activate
        ws.Cells.FormatConditions.Delete

        ws.Range("L3").Select
        Set regra = ws.Range("L3:AP3").FormatConditions.Add( _
            Type:=xlExpression, Formula1:=formulaFimDeSemana)
        regra.Interior.Color = RGB(231, 230, 230)

        Set regra = ws.Range("L4:AP" & ultimaLinha).FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""X""")
        regra.Interior.Color = RGB(225, 75, 12)

        Set faixaStatus = ws.Range("H4:H" & ultimaLinha)
        ws.Range("H4").Select

        For Each celula In listaStatus.Cells
            If Len(Trim$(CStr(celula.Value2))) > 0 _
               And celula.Interior.ColorIndex <> xlColorIndexNone Then
                Set regra = faixaStatus.FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=$H4=""" & Replace(CStr(celula.Value2), """", """""") & """")
                CopiarCor celula, regra
                regra.StopIfTrue = True
            End If
        Next celula

        'Data de finalização preenchida tem prioridade sobre o status digitado.
        If Not celulaFinalizado Is Nothing Then
            Set regra = faixaStatus.FormatConditions.Add( _
                Type:=xlExpression, Formula1:="=$K4<>""""")
            CopiarCor celulaFinalizado, regra
            regra.Font.Bold = True
            regra.StopIfTrue = True
            regra.SetFirstPriority
        End If
    Next nome

    AplicarFormatacaoStatus = avisos
End Function

Private Sub CopiarCor(ByVal origem As Range, ByVal regra As FormatCondition)
    regra.Interior.Color = origem.Interior.Color
    regra.Font.Color = origem.Font.Color
    regra.Font.Bold = origem.Font.Bold
End Sub

'Cores usadas até a versão anterior. Só pinta células da aba LISTAS que ainda
'estão sem cor; depois disso, a cor de cada status é a da própria célula.
Private Function PintarCorPadrao(ByVal celula As Range) As Boolean
    Dim fundo As Long
    Dim fonte As Long
    Dim negrito As Boolean

    Select Case UCase$(Trim$(CStr(celula.Value2)))
        Case "ABRIR ORDEM": fundo = RGB(244, 177, 131): fonte = vbBlack
        Case "SEPARAÇÃO": fundo = RGB(255, 217, 102): fonte = vbBlack
        Case "AGUARDANDO PRODUÇÃO": fundo = RGB(198, 224, 180): fonte = vbBlack
        Case "PRODUZINDO": fundo = RGB(112, 173, 71): fonte = vbWhite
        Case "FINALIZADO": fundo = RGB(91, 155, 213): fonte = vbWhite
        Case "AGUARDANDO FORMULAÇÃO": fundo = RGB(217, 217, 217): fonte = vbBlack
        Case "AGUARDANDO ENCHIMENTO": fundo = RGB(166, 166, 166): fonte = vbBlack
        Case "AGUARDANDO ENDEREÇAMENTO": fundo = RGB(244, 204, 204): fonte = vbBlack
        Case "AGUARDANDO IMPORTAÇÃO": fundo = RGB(192, 0, 0): fonte = vbWhite
        Case "FORMULANDO": fundo = RGB(255, 153, 204): fonte = vbBlack: negrito = True
        Case "ORDEM ABERTA": fundo = RGB(112, 48, 160): fonte = vbWhite: negrito = True
        Case Else
            Exit Function
    End Select

    celula.Interior.Color = fundo
    celula.Font.Color = fonte
    celula.Font.Bold = negrito
    PintarCorPadrao = True
End Function

'Fórmulas de formatação condicional criadas por VBA precisam estar no idioma
'do Excel instalado (ex.: E(...;...) em português). A célula auxiliar converte.
Private Function FormulaNoIdiomaDoExcel(ByVal formulaEmIngles As String) As String
    Dim auxiliar As Range

    Set auxiliar = ObterAba("LISTAS").Range("ZZ1")
    auxiliar.Formula = formulaEmIngles
    FormulaNoIdiomaDoExcel = auxiliar.FormulaLocal
    auxiliar.ClearContents
End Function

Public Sub AtualizarStatusFinalizado()
    Dim nomes As Variant
    Dim nome As Variant
    Dim ws As Worksheet
    Dim wsBase As Worksheet
    Dim limite As Long
    Dim linha As Long
    Dim ultimaLinhaBase As Long
    Dim posicao As Variant
    Dim op As String
    Dim situacao As String
    Dim arquivoFechado As String
    Dim fechadoNaFonte As Boolean
    Dim valorData As Variant
    Dim dataValida As Boolean

    Set wsBase = ThisWorkbook.Worksheets(ABA_BASE)
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row

    nomes = AbasOperacionais()

    For Each nome In nomes
        Set ws = ObterAba(CStr(nome))

        limite = UltimaLinhaDados(ws)

        For linha = 4 To limite
            op = NormalizarOP(ws.Cells(linha, 3).Value2)
            situacao = vbNullString
            arquivoFechado = vbNullString
            fechadoNaFonte = False

            If Len(op) > 0 And ultimaLinhaBase >= 2 Then
                posicao = Application.Match( _
                    op, wsBase.Range("A2:A" & ultimaLinhaBase), 0)

                If Not IsError(posicao) Then
                    situacao = UCase$(Trim$(CStr( _
                        wsBase.Cells(CLng(posicao) + 1, 4).Value2)))
                    arquivoFechado = Trim$(CStr( _
                        wsBase.Cells(CLng(posicao) + 1, 6).Value2))
                    fechadoNaFonte = (situacao = "FECHADA" _
                                      And Len(arquivoFechado) > 0)
                End If
            End If

            'Fora de OP_FECHADAS, o status manual não é alterado.
            If fechadoNaFonte Then
                valorData = ws.Cells(linha, 11).Value2
                dataValida = False

                If Not IsError(valorData) Then
                    If IsNumeric(valorData) Then
                        dataValida = (CDbl(valorData) > 0)
                    ElseIf IsDate(ws.Cells(linha, 11).Value) Then
                        dataValida = True
                    End If
                End If

                If dataValida Then
                    If UCase$(Trim$(CStr( _
                       ws.Cells(linha, 8).Value2))) <> "FINALIZADO" Then
                        ws.Cells(linha, 8).Value = "FINALIZADO"
                    End If
                End If
            End If
        Next linha
    Next nome
End Sub


