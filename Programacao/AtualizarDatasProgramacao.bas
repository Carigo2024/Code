Attribute VB_Name = "AtualizarDatasProgramacao"
'==============================================================================
'  AtualizarDatasProgramacao
'------------------------------------------------------------------------------
'  Objetivo:
'    Na planilha de programacao mensal, para cada O.P. digitada na coluna C:
'      - Preencher a coluna J (Data de Inicio) com a data da coluna "Emissao"
'        encontrada no ULTIMO arquivo salvo na pasta OP_ABERTAS.
'      - Preencher a coluna K (Data de Finalizacao) com a data da coluna
'        "Encerramento" (coluna D) do ULTIMO arquivo salvo na pasta OP_FECHADAS.
'      - Conferir se a "Data de Inicio" (J) confere com a data de "Emissao"
'        registrada no arquivo de OP_FECHADAS para a MESMA O.P. Caso divirjam,
'        a celula J e destacada e uma observacao e escrita na coluna L.
'
'  Observacoes de robustez (evita o erro 1004 "Erro de definicao de aplicativo
'  ou de definicao de objeto"):
'    - Verifica se as pastas existem antes de ler.
'    - Verifica se ha arquivo antes de abrir.
'    - Abre os arquivos de origem somente-leitura e sem atualizar vinculos.
'    - Localiza as colunas pelo NOME do cabecalho (nao por letra fixa), para
'      nao quebrar se o layout dos arquivos de extracao mudar de posicao.
'    - Trata O.P. nao encontrada sem interromper a rotina.
'    - Restaura as configuracoes do Excel e fecha os arquivos abertos mesmo
'      em caso de erro (rotina de limpeza no final).
'==============================================================================
Option Explicit

'--- AJUSTE AQUI CONFORME O SEU AMBIENTE -------------------------------------
Private Const PASTA_ABERTAS As String = _
    "C:\Users\CEmerson\Ultradent Products Inc\PCP e Logistica - Hard work - Documentos\CONTROL BOARD\EXTRACOES\OP_ABERTAS"
Private Const PASTA_FECHADAS As String = _
    "C:\Users\CEmerson\Ultradent Products Inc\PCP e Logistica - Hard work - Documentos\CONTROL BOARD\EXTRACOES\OP_FECHADAS"

' Nome da aba da planilha de programacao (onde estao as O.P.s na coluna C).
Private Const ABA_PROGRAMACAO As String = "Programacao"

' Linha do cabecalho e primeira linha de dados na planilha de programacao.
Private Const LINHA_CABECALHO_PROG As Long = 1
Private Const PRIMEIRA_LINHA_PROG As Long = 2

' Colunas de destino na planilha de programacao.
Private Const COL_OP_PROG As Long = 3        ' C  = O.P.
Private Const COL_INICIO_PROG As Long = 10   ' J  = Data de Inicio
Private Const COL_FIM_PROG As Long = 11      ' K  = Data de Finalizacao
Private Const COL_OBS_PROG As Long = 12      ' L  = Observacao (conferencia)

' Linha do cabecalho dentro dos arquivos de extracao (OP_ABERTAS/OP_FECHADAS).
Private Const LINHA_CABECALHO_EXTRACAO As Long = 1

' Extensao dos arquivos de extracao (ajuste se forem .xls, .csv, etc.).
Private Const EXTENSAO_EXTRACAO As String = "*.xls*"
'----------------------------------------------------------------------------


Public Sub AtualizarDatasProgramacao()

    Dim wsProg As Worksheet
    Dim wbAbertas As Workbook, wbFechadas As Workbook
    Dim wsAbertas As Worksheet, wsFechadas As Worksheet

    Dim arqAbertas As String, arqFechadas As String
    Dim colOpAb As Long, colEmissaoAb As Long
    Dim colOpFe As Long, colEmissaoFe As Long, colEncerraFe As Long

    Dim ultimaLinha As Long, i As Long
    Dim op As String
    Dim linhaAb As Long, linhaFe As Long
    Dim dtEmissaoAb As Variant, dtEncerra As Variant, dtEmissaoFe As Variant
    Dim atualizadas As Long, semAbertas As Long, semFechadas As Long, divergencias As Long

    On Error GoTo TratarErro

    '--- 1) Planilha de programacao (o proprio arquivo que roda a macro) -----
    On Error Resume Next
    Set wsProg = ThisWorkbook.Worksheets(ABA_PROGRAMACAO)
    On Error GoTo TratarErro
    If wsProg Is Nothing Then
        MsgBox "Nao encontrei a aba '" & ABA_PROGRAMACAO & "' neste arquivo." & vbNewLine & _
               "Ajuste a constante ABA_PROGRAMACAO no codigo.", vbCritical
        Exit Sub
    End If

    '--- 2) Localiza os ultimos arquivos de cada pasta ----------------------
    If Not PastaExiste(PASTA_ABERTAS) Then
        MsgBox "Pasta OP_ABERTAS nao encontrada:" & vbNewLine & PASTA_ABERTAS, vbCritical
        Exit Sub
    End If
    If Not PastaExiste(PASTA_FECHADAS) Then
        MsgBox "Pasta OP_FECHADAS nao encontrada:" & vbNewLine & PASTA_FECHADAS, vbCritical
        Exit Sub
    End If

    arqAbertas = EncontrarUltimoArquivo(PASTA_ABERTAS, EXTENSAO_EXTRACAO)
    arqFechadas = EncontrarUltimoArquivo(PASTA_FECHADAS, EXTENSAO_EXTRACAO)

    If Len(arqAbertas) = 0 Then
        MsgBox "Nenhum arquivo encontrado na pasta OP_ABERTAS.", vbExclamation
        Exit Sub
    End If
    If Len(arqFechadas) = 0 Then
        MsgBox "Nenhum arquivo encontrado na pasta OP_FECHADAS.", vbExclamation
        Exit Sub
    End If

    '--- 3) Ganho de performance --------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    '--- 4) Abre os arquivos de origem (somente leitura) --------------------
    Set wbAbertas = Workbooks.Open(Filename:=arqAbertas, ReadOnly:=True, UpdateLinks:=0)
    Set wsAbertas = wbAbertas.Worksheets(1)   ' 1a aba do arquivo de extracao

    Set wbFechadas = Workbooks.Open(Filename:=arqFechadas, ReadOnly:=True, UpdateLinks:=0)
    Set wsFechadas = wbFechadas.Worksheets(1)

    '--- 5) Descobre as colunas pelo nome do cabecalho ----------------------
    colOpAb = AcharColuna(wsAbertas, LINHA_CABECALHO_EXTRACAO, Array("OP", "O.P", "O.P.", "Ordem", "Order"))
    colEmissaoAb = AcharColuna(wsAbertas, LINHA_CABECALHO_EXTRACAO, Array("Emissao"))

    colOpFe = AcharColuna(wsFechadas, LINHA_CABECALHO_EXTRACAO, Array("OP", "O.P", "O.P.", "Ordem", "Order"))
    colEmissaoFe = AcharColuna(wsFechadas, LINHA_CABECALHO_EXTRACAO, Array("Emissao"))
    colEncerraFe = AcharColuna(wsFechadas, LINHA_CABECALHO_EXTRACAO, Array("Encerramento"))

    If colOpAb = 0 Or colEmissaoAb = 0 Then
        MsgBox "No arquivo de OP_ABERTAS nao encontrei as colunas 'OP' e/ou 'Emissao'." & vbNewLine & _
               "Arquivo: " & arqAbertas, vbCritical
        GoTo Limpar
    End If
    If colOpFe = 0 Or colEncerraFe = 0 Then
        MsgBox "No arquivo de OP_FECHADAS nao encontrei as colunas 'OP' e/ou 'Encerramento'." & vbNewLine & _
               "Arquivo: " & arqFechadas, vbCritical
        GoTo Limpar
    End If

    '--- 6) Percorre as O.P.s da planilha de programacao --------------------
    ultimaLinha = wsProg.Cells(wsProg.Rows.Count, COL_OP_PROG).End(xlUp).Row

    For i = PRIMEIRA_LINHA_PROG To ultimaLinha

        op = Trim(CStr(wsProg.Cells(i, COL_OP_PROG).Value))
        If Len(op) = 0 Then GoTo ProximaLinha   ' ignora linhas sem O.P.

        ' Limpa observacao anterior desta linha
        wsProg.Cells(i, COL_OBS_PROG).ClearContents
        wsProg.Cells(i, COL_INICIO_PROG).Interior.Pattern = xlNone

        ' ---- Data de Inicio (J) = Emissao no arquivo de OP_ABERTAS ----
        linhaAb = ProcurarLinhaPorOP(wsAbertas, colOpAb, LINHA_CABECALHO_EXTRACAO, op)
        If linhaAb > 0 Then
            dtEmissaoAb = ConverterData(wsAbertas.Cells(linhaAb, colEmissaoAb).Value)
            If IsDate(dtEmissaoAb) Then
                wsProg.Cells(i, COL_INICIO_PROG).Value = CDate(dtEmissaoAb)
                wsProg.Cells(i, COL_INICIO_PROG).NumberFormat = "yyyy-mm-dd"
                atualizadas = atualizadas + 1
            End If
        Else
            semAbertas = semAbertas + 1
            wsProg.Cells(i, COL_OBS_PROG).Value = "O.P. nao encontrada em OP_ABERTAS"
        End If

        ' ---- Data de Finalizacao (K) = Encerramento no arquivo de OP_FECHADAS ----
        linhaFe = ProcurarLinhaPorOP(wsFechadas, colOpFe, LINHA_CABECALHO_EXTRACAO, op)
        If linhaFe > 0 Then
            dtEncerra = ConverterData(wsFechadas.Cells(linhaFe, colEncerraFe).Value)
            If IsDate(dtEncerra) Then
                wsProg.Cells(i, COL_FIM_PROG).Value = CDate(dtEncerra)
                wsProg.Cells(i, COL_FIM_PROG).NumberFormat = "yyyy-mm-dd"
            End If

            ' ---- Conferencia: Emissao (OP_FECHADAS) x Data de Inicio (J) ----
            If colEmissaoFe > 0 And linhaAb > 0 Then
                dtEmissaoFe = ConverterData(wsFechadas.Cells(linhaFe, colEmissaoFe).Value)
                If IsDate(dtEmissaoFe) And IsDate(dtEmissaoAb) Then
                    If CLng(CDate(dtEmissaoFe)) <> CLng(CDate(dtEmissaoAb)) Then
                        divergencias = divergencias + 1
                        wsProg.Cells(i, COL_INICIO_PROG).Interior.Color = RGB(255, 199, 206) ' vermelho claro
                        wsProg.Cells(i, COL_OBS_PROG).Value = _
                            "Emissao divergente: ABERTAS=" & Format$(CDate(dtEmissaoAb), "yyyy-mm-dd") & _
                            " x FECHADAS=" & Format$(CDate(dtEmissaoFe), "yyyy-mm-dd")
                    End If
                End If
            End If
        Else
            semFechadas = semFechadas + 1
            If Len(wsProg.Cells(i, COL_OBS_PROG).Value) = 0 Then
                wsProg.Cells(i, COL_OBS_PROG).Value = "O.P. nao encontrada em OP_FECHADAS"
            End If
        End If

ProximaLinha:
    Next i

Limpar:
    '--- 7) Fecha origens e restaura o Excel --------------------------------
    On Error Resume Next
    If Not wbAbertas Is Nothing Then wbAbertas.Close SaveChanges:=False
    If Not wbFechadas Is Nothing Then wbFechadas.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    On Error GoTo 0

    If Not wsProg Is Nothing Then
        MsgBox "Concluido." & vbNewLine & _
               "Datas de inicio preenchidas: " & atualizadas & vbNewLine & _
               "O.P. sem registro em ABERTAS: " & semAbertas & vbNewLine & _
               "O.P. sem registro em FECHADAS: " & semFechadas & vbNewLine & _
               "Divergencias de Emissao: " & divergencias, vbInformation
    End If
    Exit Sub

TratarErro:
    MsgBox "Ocorreu um erro (" & Err.Number & "): " & Err.Description, vbCritical
    Resume Limpar

End Sub


'==============================================================================
'  FUNCOES AUXILIARES
'==============================================================================

' Retorna o caminho completo do arquivo mais recente (por data de modificacao)
' dentro de 'pasta' que corresponda ao 'padrao' (ex.: "*.xls*").
Private Function EncontrarUltimoArquivo(ByVal pasta As String, ByVal padrao As String) As String
    Dim nome As String, caminho As String
    Dim maisRecente As String, dtMaisRecente As Date, dtArquivo As Date

    If Right$(pasta, 1) <> "\" Then pasta = pasta & "\"

    nome = Dir(pasta & padrao, vbNormal)
    Do While Len(nome) > 0
        ' Ignora arquivos temporarios do Excel (comecam com ~$)
        If Left$(nome, 2) <> "~$" Then
            caminho = pasta & nome
            dtArquivo = FileDateTime(caminho)
            If Len(maisRecente) = 0 Or dtArquivo > dtMaisRecente Then
                maisRecente = caminho
                dtMaisRecente = dtArquivo
            End If
        End If
        nome = Dir
    Loop

    EncontrarUltimoArquivo = maisRecente
End Function


' Verifica se uma pasta existe.
Private Function PastaExiste(ByVal caminho As String) As Boolean
    On Error Resume Next
    PastaExiste = (Len(Dir(caminho, vbDirectory)) > 0)
    On Error GoTo 0
End Function


' Procura, na 'linhaCabecalho' de 'ws', a primeira coluna cujo cabecalho
' (normalizado, sem acento e sem diferenciar maiuscula/minuscula) bata com
' algum dos nomes candidatos. Retorna o indice da coluna ou 0 se nao achar.
Private Function AcharColuna(ByVal ws As Worksheet, ByVal linhaCabecalho As Long, _
                             ByVal candidatos As Variant) As Long
    Dim ultimaCol As Long, c As Long, k As Long
    Dim cab As String

    ultimaCol = ws.Cells(linhaCabecalho, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To ultimaCol
        cab = Normalizar(CStr(ws.Cells(linhaCabecalho, c).Value))
        For k = LBound(candidatos) To UBound(candidatos)
            If cab = Normalizar(CStr(candidatos(k))) Then
                AcharColuna = c
                Exit Function
            End If
        Next k
    Next c

    AcharColuna = 0
End Function


' Procura a linha em que a coluna 'colOp' de 'ws' contem o valor 'op'.
' Retorna o numero da linha ou 0 se nao encontrar.
Private Function ProcurarLinhaPorOP(ByVal ws As Worksheet, ByVal colOp As Long, _
                                    ByVal linhaCabecalho As Long, ByVal op As String) As Long
    Dim rng As Range, achou As Range
    Dim ultimaLinha As Long

    ultimaLinha = ws.Cells(ws.Rows.Count, colOp).End(xlUp).Row
    If ultimaLinha <= linhaCabecalho Then Exit Function

    Set rng = ws.Range(ws.Cells(linhaCabecalho + 1, colOp), ws.Cells(ultimaLinha, colOp))

    Set achou = rng.Find(What:=op, LookIn:=xlValues, LookAt:=xlWhole, _
                         MatchCase:=False, SearchFormat:=False)
    If Not achou Is Nothing Then
        ProcurarLinhaPorOP = achou.Row
    Else
        ProcurarLinhaPorOP = 0
    End If
End Function


' Converte um valor (data real ou texto "aaaa-mm-dd") para uma data.
' Retorna a data como Variant; se nao for possivel, retorna vazio.
Private Function ConverterData(ByVal v As Variant) As Variant
    Dim s As String
    Dim a As Integer, m As Integer, d As Integer

    If IsDate(v) Then
        ConverterData = CDate(v)
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        ConverterData = Empty
        Exit Function
    End If

    ' Tenta o formato explicito aaaa-mm-dd (evita ambiguidade de local do Excel)
    s = Replace(s, "/", "-")
    If Len(s) >= 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        If IsNumeric(Left$(s, 4)) And IsNumeric(Mid$(s, 6, 2)) And IsNumeric(Mid$(s, 9, 2)) Then
            a = CInt(Left$(s, 4))
            m = CInt(Mid$(s, 6, 2))
            d = CInt(Mid$(s, 9, 2))
            On Error Resume Next
            ConverterData = DateSerial(a, m, d)
            On Error GoTo 0
            Exit Function
        End If
    End If

    ' Ultimo recurso: deixa o Excel tentar interpretar
    On Error Resume Next
    ConverterData = CDate(s)
    On Error GoTo 0
End Function


' Normaliza texto: remove espacos, deixa minusculo e troca acentos comuns.
Private Function Normalizar(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, "á", "a"): s = Replace(s, "à", "a"): s = Replace(s, "ã", "a"): s = Replace(s, "â", "a")
    s = Replace(s, "é", "e"): s = Replace(s, "ê", "e")
    s = Replace(s, "í", "i")
    s = Replace(s, "ó", "o"): s = Replace(s, "õ", "o"): s = Replace(s, "ô", "o")
    s = Replace(s, "ú", "u")
    s = Replace(s, "ç", "c")
    Normalizar = s
End Function
