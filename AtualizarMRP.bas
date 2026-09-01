Attribute VB_Name = "AtualizarMRP"
'=====================================================================
' AUTOMAÇÃO DE ATUALIZAÇÃO DA ANÁLISE MRP - PROTHEUS  (v4)
' TUDO sobre a MÉDIA MENSAL. Regras:
'   - Saída = "- Saídas" + "- Saída Estrutura" (unificada).
'   - Necessidade e Saída pela MÉDIA dos 12 meses.
'   - Compra/Produção sugeridas = necessidade MÉDIA MENSAL (com lote mín/econ).
'   - Necessidade média líquida = média mensal cheia (recebimento só afeta status).
'   - Produzido = apenas Tipo "PA"; demais = Comprado.
'   - Urgente se (Estoque Atual + Estoque Segurança) em dias < Lead Time.
' Analise_MRP tem 20 colunas (A..T), sem colunas de total.
'
' Histórico v4 (revisão de código):
'   - Corrigida codificação mista (ANSI/UTF-8): os rótulos "- Saídas" e
'     "- Saída Estrutura" agora casam com o relatório do Protheus.
'   - ValN reescrita para parsing numérico independente de locale (Val()).
'   - Cancelamento do seletor de arquivo detectado por VarType (robusto).
'   - Verificação de existência das abas obrigatórias com mensagem clara.
' IMPORTANTE: salvar/importar este arquivo em ANSI (Windows-1252) no VBE.
'=====================================================================
Option Explicit

Const ABA_RELATORIO As String = "Resultados"
Const ABA_ANALISE As String = "Analise_MRP"
Const ABA_RESUMO As String = "Resumo_Executivo"
Const ABA_BASE As String = "Base_Original"
Const NPER As Integer = 12
Const PASTA_PADRAO As String = "C:\MRP\"

Sub AtualizarMRPV1()
    Dim caminhoArq As Variant
    Dim wbNovo As Workbook, wsRel As Worksheet
    Dim t0 As Single: t0 = Timer

    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' Valida abas obrigatórias antes de qualquer processamento
    If Not AbaExiste(ABA_ANALISE) Then
        MsgBox "Aba """ & ABA_ANALISE & """ não encontrada nesta pasta de trabalho.", vbCritical
        GoTo Limpeza
    End If

    caminhoArq = SelecionarArquivo()
    If VarType(caminhoArq) = vbBoolean Then
        MsgBox "Atualização cancelada.", vbInformation
        GoTo Limpeza
    End If

    Set wbNovo = Workbooks.Open(caminhoArq, ReadOnly:=True)
    On Error Resume Next
    Set wsRel = wbNovo.Worksheets(ABA_RELATORIO)
    On Error GoTo TrataErro
    If wsRel Is Nothing Then Set wsRel = wbNovo.Worksheets(1)

    CopiarBaseOriginal wsRel

    Dim periodos(1 To NPER) As Variant
    LerPeriodos wsRel, periodos

    Dim dados() As Variant, nProd As Long
    nProd = ProcessarBlocos(wsRel, dados, periodos)

    wbNovo.Close SaveChanges:=False: Set wbNovo = Nothing
    If nProd = 0 Then MsgBox "Nenhum produto encontrado.", vbExclamation: GoTo Limpeza

    EscreverAnalise dados, nProd

    ' Recalcula o resumo executivo apenas se a aba existir
    If AbaExiste(ABA_RESUMO) Then ThisWorkbook.Worksheets(ABA_RESUMO).Calculate

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    MsgBox "Análise MRP atualizada!" & vbCrLf & nProd & " produtos em " & Format(Timer - t0, "0.0") & "s.", vbInformation

Limpeza:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Exit Sub
TrataErro:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    If Not wbNovo Is Nothing Then wbNovo.Close SaveChanges:=False
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Function AbaExiste(ByVal nome As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    AbaExiste = Not (ws Is Nothing)
End Function

Private Function SelecionarArquivo() As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o novo relatório do MRP (Protheus)"
        .Filters.Clear: .Filters.Add "Excel", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False
        If Dir(PASTA_PADRAO, vbDirectory) <> "" Then .InitialFileName = PASTA_PADRAO
        If .Show = -1 Then SelecionarArquivo = .SelectedItems(1) Else SelecionarArquivo = False
    End With
End Function

Private Sub CopiarBaseOriginal(wsRel As Worksheet)
    Dim wsDest As Worksheet
    On Error Resume Next
    Set wsDest = ThisWorkbook.Worksheets(ABA_BASE)
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsDest.Name = ABA_BASE
    End If
    wsDest.Cells.Clear
    Dim ur As Range: Set ur = wsRel.UsedRange
    wsDest.Range(ur.Address).Value = ur.Value
End Sub

Private Sub LerPeriodos(wsRel As Worksheet, periodos() As Variant)
    Dim linhaPer As Long, c As Integer
    linhaPer = AcharLinhaPeriodo(wsRel)
    For c = 1 To NPER: periodos(c) = wsRel.Cells(linhaPer, c + 1).Value: Next c
End Sub

Private Function AcharLinhaPeriodo(wsRel As Worksheet) As Long
    Dim r As Long
    For r = 1 To 20
        If InStr(1, CStr(wsRel.Cells(r, 1).Value), "Per", vbTextCompare) = 1 Then AcharLinhaPeriodo = r: Exit Function
    Next r
    AcharLinhaPeriodo = 3
End Function

Private Function ProcessarBlocos(wsRel As Worksheet, dados() As Variant, periodos() As Variant) As Long
    Dim ultLinha As Long, r As Long
    ultLinha = wsRel.Cells(wsRel.Rows.Count, 1).End(xlUp).Row
    Dim m As Variant
    m = wsRel.Range(wsRel.Cells(1, 1), wsRel.Cells(ultLinha, 14)).Value

    Dim nProd As Long: nProd = 0
    For r = 1 To ultLinha
        If Trim(CStr(m(r, 1))) = "Produto" Then nProd = nProd + 1
    Next r
    If nProd = 0 Then ProcessarBlocos = 0: Exit Function

    ' 22 campos (20 de saída + 2 auxiliares de ordenação)
    ReDim dados(1 To nProd, 1 To 22)
    Dim idx As Long: idx = 0
    Dim k As Integer
    r = 1
    Do While r <= ultLinha
        If Trim(CStr(m(r, 1))) = "Produto" Then
            Dim rMestre As Long: rMestre = r + 1
            idx = idx + 1

            Dim codigo As String, descr As String, tipo As String, um As String
            Dim loteMin As Double, loteEcon As Double, valorUnit As Double
            Dim leadTime As Double, estSeg As Double
            codigo = Trim(CStr(m(rMestre, 1))): descr = CStr(m(rMestre, 2)): tipo = Trim(CStr(m(rMestre, 3)))
            leadTime = ValN(m(rMestre, 5)): loteMin = ValN(m(rMestre, 9)): loteEcon = ValN(m(rMestre, 10))
            estSeg = ValN(m(rMestre, 11)): um = CStr(m(rMestre, 13)): valorUnit = ValN(m(rMestre, 14))

            Dim estoqueAtual As Double, recebPlan As Double, saidaTotal As Double
            Dim necTotal As Double, saldoMin As Double, idxNec As Integer
            estoqueAtual = 0: recebPlan = 0: saidaTotal = 0: necTotal = 0: saldoMin = 1E+15: idxNec = 0

            Dim rr As Long: rr = rMestre + 1
            Do While rr <= ultLinha
                If Trim(CStr(m(rr, 1))) = "Produto" Then Exit Do
                Dim rot As String: rot = Trim(CStr(m(rr, 1)))
                Select Case rot
                    Case "+ Estoque": estoqueAtual = ValN(m(rr, 2))
                    Case "+ Entradas"
                        For k = 1 To NPER: recebPlan = recebPlan + ValN(m(rr, k + 1)): Next k
                    Case "- Saídas"
                        For k = 1 To NPER: saidaTotal = saidaTotal + ValN(m(rr, k + 1)): Next k
                    Case "- Saída Estrutura"
                        For k = 1 To NPER: saidaTotal = saidaTotal + ValN(m(rr, k + 1)): Next k
                    Case "Saldo Final"
                        For k = 1 To NPER
                            If ValN(m(rr, k + 1)) < saldoMin Then saldoMin = ValN(m(rr, k + 1))
                        Next k
                    Case "Necessidade"
                        For k = 1 To NPER
                            necTotal = necTotal + ValN(m(rr, k + 1))
                            If ValN(m(rr, k + 1)) > 0 And idxNec = 0 Then idxNec = k
                        Next k
                End Select
                rr = rr + 1
            Loop
            If saldoMin = 1E+15 Then saldoMin = 0

            ' MÉDIAS mensais
            Dim necMedia As Double, saidaMedia As Double
            necMedia = necTotal / NPER
            saidaMedia = saidaTotal / NPER

            Dim isProduzido As Boolean: isProduzido = (tipo = "PA")
            Dim natureza As String
            If tipo = "" Then
                natureza = "Verificar"
            ElseIf isProduzido Then
                natureza = "Produzir"
            Else
                natureza = "Comprar"
            End If

            Dim consumoDia As Double, coberturaDias As Double
            If saidaMedia > 0 Then
                consumoDia = saidaMedia / 30#: coberturaDias = (estoqueAtual + estSeg) / consumoDia
            Else
                consumoDia = 0: coberturaDias = 9999
            End If

            Dim status As String, prioridade As String, obs As String: obs = ""
            If tipo = "" Then
                status = "Verificar cadastro ou parâmetro": prioridade = "Verificar": obs = "Tipo de item ausente"
            ElseIf necMedia <= 0 And saldoMin >= 0 Then
                Dim cob As Double: cob = estoqueAtual + recebPlan
                If saidaMedia > 0 And cob > saidaTotal * 3 Then
                    status = "Excesso de estoque": prioridade = "Baixa": obs = "Estoque/recebimento muito acima da demanda"
                ElseIf saidaMedia = 0 And estoqueAtual > 0 Then
                    status = "Sem ação imediata": prioridade = "Baixa": obs = "Sem demanda no horizonte"
                Else
                    status = "Coberto por estoque": prioridade = "Baixa"
                End If
            Else
                Dim urgente As Boolean: urgente = (coberturaDias < leadTime)
                If recebPlan >= necTotal And recebPlan > 0 Then
                    status = "Coberto por recebimento planejado": prioridade = "Média": obs = "Recebimento planejado cobre a necessidade"
                ElseIf isProduzido Then
                    If urgente Then
                        status = "Produzir urgente": prioridade = "Alta"
                    Else
                        status = "Produzir": prioridade = "Média"
                    End If
                Else
                    If urgente Then
                        status = "Comprar urgente": prioridade = "Alta"
                    Else
                        status = "Comprar": prioridade = "Média"
                    End If
                End If
                If urgente Then obs = "Cobertura " & Format(coberturaDias, "0") & "d < lead time " & Format(leadTime, "0") & "d"
            End If

            ' NECESSIDADE MÉDIA LÍQUIDA = média mensal cheia; qtd sugerida com lote
            Dim necMediaLiq As Double: necMediaLiq = IIf(necMedia > 0, necMedia, 0)
            Dim qtComprar As Double, qtProduzir As Double: qtComprar = 0: qtProduzir = 0
            If necMediaLiq > 0 Then
                If isProduzido Then
                    qtProduzir = necMediaLiq
                    If loteEcon > 0 And qtProduzir < loteEcon Then
                        qtProduzir = loteEcon
                    ElseIf loteMin > 0 And qtProduzir < loteMin Then
                        qtProduzir = loteMin
                    End If
                Else
                    qtComprar = necMediaLiq
                    If loteEcon > 0 And qtComprar < loteEcon Then
                        qtComprar = loteEcon
                    ElseIf loteMin > 0 And qtComprar < loteMin Then
                        qtComprar = loteMin
                    End If
                End If
            End If

            Dim dataCritica As Variant
            If idxNec >= 1 Then dataCritica = periodos(idxNec) Else dataCritica = ""
            Dim coberturaOut As Variant
            If coberturaDias < 9999 Then coberturaOut = Round(coberturaDias, 0) Else coberturaOut = ""
            Dim valorNec As Double: valorNec = necMediaLiq * valorUnit

            ' Gravar (20 colunas + 2 auxiliares)
            dados(idx, 1) = codigo: dados(idx, 2) = descr: dados(idx, 3) = tipo
            dados(idx, 4) = natureza: dados(idx, 5) = um
            dados(idx, 6) = estoqueAtual: dados(idx, 7) = estSeg: dados(idx, 8) = recebPlan
            dados(idx, 9) = Round(saidaMedia, 1): dados(idx, 10) = Round(necMedia, 1)
            dados(idx, 11) = saldoMin: dados(idx, 12) = leadTime: dados(idx, 13) = coberturaOut
            dados(idx, 14) = Round(qtComprar, 1): dados(idx, 15) = Round(qtProduzir, 1)
            dados(idx, 16) = dataCritica: dados(idx, 17) = status: dados(idx, 18) = prioridade
            dados(idx, 19) = Round(valorNec, 2): dados(idx, 20) = obs
            dados(idx, 21) = OrdemPrioridade(prioridade): dados(idx, 22) = valorNec

            r = rr
        Else
            r = r + 1
        End If
    Loop
    ProcessarBlocos = idx
End Function

Private Sub EscreverAnalise(dados() As Variant, nProd As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(ABA_ANALISE)
    OrdenarDados dados, nProd
    Dim ultAnt As Long: ultAnt = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultAnt >= 3 Then ws.Range("A3:T" & ultAnt).ClearContents

    Dim saida() As Variant: ReDim saida(1 To nProd, 1 To 20)
    Dim i As Long, j As Integer
    For i = 1 To nProd
        For j = 1 To 20: saida(i, j) = dados(i, j): Next j
    Next i
    ws.Range("A3").Resize(nProd, 20).Value = saida

    ws.Range("F3:O" & (nProd + 2)).NumberFormat = "#,##0"
    ws.Range("S3:S" & (nProd + 2)).NumberFormat = "#,##0.00"
    ws.Range("P3:P" & (nProd + 2)).NumberFormat = "DD/MM/YYYY"
    On Error Resume Next
    ws.AutoFilterMode = False
    ws.Range("A2:T" & (nProd + 2)).AutoFilter
    On Error GoTo 0
End Sub

Private Sub OrdenarDados(dados() As Variant, nProd As Long)
    Dim i As Long, j As Long, k As Integer, troca As Boolean, tmp As Variant
    For i = 1 To nProd - 1
        For j = 1 To nProd - i
            troca = False
            If dados(j, 21) > dados(j + 1, 21) Then
                troca = True
            ElseIf dados(j, 21) = dados(j + 1, 21) Then
                If ValN(dados(j, 22)) < ValN(dados(j + 1, 22)) Then troca = True
            End If
            If troca Then
                For k = 1 To 22
                    tmp = dados(j, k): dados(j, k) = dados(j + 1, k): dados(j + 1, k) = tmp
                Next k
            End If
        Next j
    Next i
End Sub

' Converte um valor de célula em Double de forma robusta e independente de locale.
' - Variant numérico: coerção direta.
' - Texto em formato brasileiro ("1.234,56"): ponto = milhar, vírgula = decimal.
'   Após normalizar para ".", usa Val(), que sempre trata "." como separador
'   decimal (evita o erro de CDbl em Windows configurado como pt-BR).
Private Function ValN(ByVal v As Variant) As Double
    On Error GoTo Falha
    If IsEmpty(v) Or IsNull(v) Then ValN = 0: Exit Function
    If VarType(v) <> vbString And IsNumeric(v) Then ValN = CDbl(v): Exit Function

    Dim s As String: s = Trim$(CStr(v))
    If Len(s) = 0 Then ValN = 0: Exit Function
    s = Replace(s, ".", "")   ' remove separador de milhar
    s = Replace(s, ",", ".")  ' vírgula decimal -> ponto
    ValN = Val(s)
    Exit Function
Falha:
    ValN = 0
End Function

Private Function OrdemPrioridade(ByVal p As String) As Integer
    Select Case p
        Case "Alta": OrdemPrioridade = 0
        Case "Média": OrdemPrioridade = 1
        Case "Baixa": OrdemPrioridade = 2
        Case Else: OrdemPrioridade = 3
    End Select
End Function
