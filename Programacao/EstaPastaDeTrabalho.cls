Private Sub Workbook_Open()
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
End Sub

Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
    AtualizarStatusDaAba Sh
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim areaAlterada As Range
    Dim celula As Range
    Dim linhas As Object
    Dim chave As Variant
    Dim eventosAnteriores As Boolean

    If Not AbaOperacional(Sh.Name) Then Exit Sub

    Set areaAlterada = Intersect(Target, Sh.Range("C:D"))
    If areaAlterada Is Nothing Then Exit Sub

    eventosAnteriores = Application.EnableEvents
    On Error GoTo Saida
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic

    Set linhas = CreateObject("Scripting.Dictionary")

    For Each celula In areaAlterada.Cells
        If celula.Row >= 4 Then linhas(CStr(celula.Row)) = True
    Next celula

    For Each chave In linhas.Keys
        If Not Sh.Cells(CLng(chave), 5).HasFormula _
           And IsEmpty(Sh.Cells(CLng(chave), 5).Value2) Then
            Sh.Cells(CLng(chave), 5).Formula2R1C1 = FormulaDescricao()
        End If

        Sh.Cells(CLng(chave), 7).Formula2R1C1 = FormulaQtd()
        Sh.Cells(CLng(chave), 10).Formula2R1C1 = FormulaEmissao()
        Sh.Cells(CLng(chave), 11).Formula2R1C1 = FormulaEncerramento()
        Sh.Cells(CLng(chave), 7).NumberFormat = "#,##0"
        Sh.Range(Sh.Cells(CLng(chave), 10), _
                 Sh.Cells(CLng(chave), 11)).NumberFormat = "dd/mm/yyyy"
    Next chave

    Sh.Calculate
    AtualizarStatusDaAba Sh

Saida:
    On Error Resume Next
    Application.EnableEvents = eventosAnteriores
    On Error GoTo 0
End Sub

Private Sub AtualizarStatusDaAba(ByVal Sh As Object)
    Dim limite As Long
    Dim linha As Long
    Dim wsBase As Worksheet
    Dim ultimaLinhaBase As Long
    Dim posicao As Variant
    Dim op As String
    Dim situacao As String
    Dim arquivoFechado As String
    Dim fechadoNaFonte As Boolean
    Dim valorData As Variant
    Dim dataValida As Boolean
    Dim eventosAnteriores As Boolean

    If Not AbaOperacional(Sh.Name) Then Exit Sub

    eventosAnteriores = Application.EnableEvents
    On Error GoTo Saida

    limite = Application.Max(4, _
        Sh.Cells(Sh.Rows.Count, 3).End(xlUp).Row, _
        Sh.Cells(Sh.Rows.Count, 4).End(xlUp).Row)

    Set wsBase = ThisWorkbook.Worksheets("_BASE_OP")
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row
    Application.EnableEvents = False

    For linha = 4 To limite
        op = NormalizarOPEvento(Sh.Cells(linha, 3).Value2)
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

        If fechadoNaFonte Then
            valorData = Sh.Cells(linha, 11).Value2
            dataValida = False

            If Not IsError(valorData) Then
                If IsNumeric(valorData) Then
                    dataValida = (CDbl(valorData) > 0)
                ElseIf IsDate(Sh.Cells(linha, 11).Value) Then
                    dataValida = True
                End If
            End If

            If dataValida Then
                If UCase$(Trim$(CStr( _
                   Sh.Cells(linha, 8).Value2))) <> "FINALIZADO" Then
                    Sh.Cells(linha, 8).Value = "FINALIZADO"
                End If
            End If
        End If
    Next linha

Saida:
    On Error Resume Next
    Application.EnableEvents = eventosAnteriores
    On Error GoTo 0
End Sub

Private Function AbaOperacional(ByVal nomeAba As String) As Boolean
    Dim nomes As Variant
    Dim nome As Variant

    'A lista de abas fica só no Módulo1 (função AbasOperacionais).
    nomes = AbasOperacionais()

    For Each nome In nomes
        If StrComp(CStr(nome), nomeAba, vbTextCompare) = 0 Then
            AbaOperacional = True
            Exit Function
        End If
    Next nome
End Function

Private Function NormalizarOPEvento(ByVal valor As Variant) As String
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
        NormalizarOPEvento = Right$("000000" & digitos, 6)
    Else
        NormalizarOPEvento = "0" & Left$(digitos, 5)
    End If
End Function
