Attribute VB_Name = "InstalarAutomacao"
'==============================================================================
'  InstalarAutomacao  -  VERSAO REVISADA E CORRIGIDA
'------------------------------------------------------------------------------
'  Le os ultimos arquivos de OP_ABERTAS e OP_FECHADAS, consolida na aba
'  oculta _BASE_OP e atualiza status/validacoes da Programacao.
'
'  Principais correcoes em relacao ao codigo original (ver REVISAO_CODIGO.md):
'   1. PLANILHAS PROTEGIDAS: qualquer escrita/validacao em aba protegida gera o
'      erro 1004 "Erro de definicao de aplicativo ou de definicao de objeto"
'      (o erro da sua imagem). Agora as abas sao desprotegidas antes e
'      reprotegidas depois (helpers Desproteger/Reproteger).
'   2. ABAS INEXISTENTES: acessar Worksheets("Nome") de uma aba que nao existe
'      gera erro 9. Agora ha verificacao previa (PlanilhaExiste) e a aba
'      ausente e apenas ignorada, sem derrubar a rotina.
'   3. MENSAGENS DE ERRO: passam a informar Err.Number e a ETAPA em que o erro
'      ocorreu, para diagnostico (a mensagem original nao dizia onde falhou).
'   4. VALORES DE ERRO na coluna K tratados antes de CStr/CDbl (evita erro 13).
'   5. Verificacao explicita de existencia das pastas de extracao.
'==============================================================================
Option Explicit

'--- AJUSTE AQUI CONFORME O SEU AMBIENTE -------------------------------------
Private Const PASTA_ABERTAS As String = "C:\Users\CEmerson\Ultradent Products Inc\PCP e Logística - Hard work - Documentos\CONTROL BOARD\EXTRACOES\OP_ABERTAS"
Private Const PASTA_FECHADAS As String = "C:\Users\CEmerson\Ultradent Products Inc\PCP e Logística - Hard work - Documentos\CONTROL BOARD\EXTRACOES\OP_FECHADAS"
Private Const NOME_BASE As String = "_BASE_OP"

' Se as abas da Programacao tiverem senha de protecao, informe-a aqui.
' Deixe "" se a protecao nao usa senha.
Private Const SENHA_PROTECAO As String = ""
'----------------------------------------------------------------------------

' Guarda a etapa atual para diagnostico em caso de erro.
Private gEtapa As String

'==============================================================================
Public Sub InstalarAutomacao()
    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Instalando automação da Programação..."

    gEtapa = "Remover consultas incompletas"
    RemoverConsultasIncompletas

    gEtapa = "Configurar lista de status"
    ConfigurarListaStatus

    gEtapa = "Atualizar programação"
    AtualizarProgramacao

    gEtapa = "Ocultar aba de instalação"
    If PlanilhaExiste("INSTALAR_AUTOMACAO") Then
        ThisWorkbook.Worksheets("INSTALAR_AUTOMACAO").Visible = xlSheetVeryHidden
    End If

    RestaurarApp
    MsgBox "Automação instalada e base atualizada com sucesso.", vbInformation
    Exit Sub

TrataErro:
    RestaurarApp
    MsgBox "Não foi possível concluir a instalação." & vbNewLine & _
           "Etapa: " & gEtapa & vbNewLine & _
           "Erro " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

'==============================================================================
Public Sub AtualizarProgramacao()
    Dim abertas As Object
    Dim fechadas As Object
    Dim arquivosAbertas As Long
    Dim arquivosFechadas As Long
    Dim segurancaAnterior As Long

    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Lendo arquivos de OP abertas e fechadas..."

    ' Verifica pastas antes de tentar ler (mensagem mais clara).
    gEtapa = "Verificar pastas de extração"
    If Not PastaExiste(PASTA_ABERTAS) Then
        Err.Raise vbObjectError + 1000, , "Pasta OP_ABERTAS não encontrada: " & PASTA_ABERTAS
    End If
    If Not PastaExiste(PASTA_FECHADAS) Then
        Err.Raise vbObjectError + 1000, , "Pasta OP_FECHADAS não encontrada: " & PASTA_FECHADAS
    End If

    Set abertas = CreateObject("Scripting.Dictionary")
    Set fechadas = CreateObject("Scripting.Dictionary")
    abertas.CompareMode = vbTextCompare
    fechadas.CompareMode = vbTextCompare

    segurancaAnterior = Application.AutomationSecurity
    Application.AutomationSecurity = 3   ' msoAutomationSecurityForceDisable (bloqueia macros dos arquivos lidos)

    gEtapa = "Ler OP_ABERTAS"
    LerPasta PASTA_ABERTAS, False, abertas, arquivosAbertas
    gEtapa = "Ler OP_FECHADAS"
    LerPasta PASTA_FECHADAS, True, fechadas, arquivosFechadas

    Application.AutomationSecurity = segurancaAnterior

    If arquivosAbertas = 0 Then
        Err.Raise vbObjectError + 1001, , "Nenhum arquivo válido foi encontrado em OP_ABERTAS."
    End If
    If arquivosFechadas = 0 Then
        Err.Raise vbObjectError + 1002, , "Nenhum arquivo válido foi encontrado em OP_FECHADAS."
    End If

    gEtapa = "Gravar base consolidada"
    GravarBase abertas, fechadas
    Application.CalculateFull

    gEtapa = "Atualizar status FINALIZADO"
    AtualizarStatusFinalizado

Saida:
    On Error Resume Next
    Application.AutomationSecurity = segurancaAnterior
    On Error GoTo 0
    RestaurarApp
    Exit Sub

TrataErro:
    MsgBox "Falha ao atualizar a Programação." & vbNewLine & _
           "Etapa: " & gEtapa & vbNewLine & _
           "Erro " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Saida
End Sub

'==============================================================================
Public Sub Auto_Open()
    ' Nao interrompe a abertura do arquivo caso as pastas estejam indisponiveis.
    On Error Resume Next
    AtualizarProgramacao
    On Error GoTo 0
End Sub

'==============================================================================
Private Sub LerPasta(ByVal pasta As String, ByVal arquivoFechado As Boolean, ByVal dados As Object, ByRef arquivosLidos As Long)
    Dim nome As String
    Dim caminho As String
    Dim extensao As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim valores As Variant
    Dim i As Long
    Dim op As String
    Dim emissao As Variant
    Dim encerramento As Variant
    Dim dataArquivo As Date
    Dim registro As Variant

    nome = Dir$(pasta & "\*.*")
    Do While Len(nome) > 0
        extensao = LCase$(Mid$(nome, InStrRev(nome, ".")))
        If Left$(nome, 2) <> "~$" And _
           (extensao = ".xlsx" Or extensao = ".xlsm" Or extensao = ".xlsb") Then

            caminho = pasta & "\" & nome
            Application.StatusBar = "Lendo " & nome & "..."

            Set wb = Nothing
            On Error Resume Next
            Set wb = Workbooks.Open( _
                Filename:=caminho, _
                UpdateLinks:=0, _
                ReadOnly:=True, _
                AddToMru:=False, _
                Notify:=False)
            On Error GoTo 0

            If Not wb Is Nothing Then
                Set ws = LocalizarExport(wb)
                If Not ws Is Nothing Then
                    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
                    If ultimaLinha >= 2 Then
                        valores = ws.Range("B1:D" & ultimaLinha).Value2
                        dataArquivo = ObterDataArquivo(nome, caminho)

                        For i = 2 To UBound(valores, 1)
                            op = NormalizarOP(valores(i, 1))
                            If Len(op) > 0 Then
                                emissao = valores(i, 2)

                                If arquivoFechado Then
                                    encerramento = valores(i, 3)
                                    If dados.Exists(op) Then
                                        registro = dados(op)
                                        If CDbl(dataArquivo) >= CDbl(registro(2)) Then
                                            dados(op) = Array(emissao, encerramento, dataArquivo, nome)
                                        End If
                                    Else
                                        dados.Add op, Array(emissao, encerramento, dataArquivo, nome)
                                    End If
                                Else
                                    If dados.Exists(op) Then
                                        registro = dados(op)
                                        If CDbl(dataArquivo) >= CDbl(registro(1)) Then
                                            dados(op) = Array(emissao, dataArquivo, nome)
                                        End If
                                    Else
                                        dados.Add op, Array(emissao, dataArquivo, nome)
                                    End If
                                End If
                            End If
                        Next i
                    End If
                    arquivosLidos = arquivosLidos + 1
                End If
                wb.Close SaveChanges:=False
            End If
        End If
        nome = Dir$()
    Loop
End Sub

'==============================================================================
Private Function LocalizarExport(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If UCase$(Trim$(ws.Name)) = "EXPORT" Then
            Set LocalizarExport = ws
            Exit Function
        End If
    Next ws
End Function

'==============================================================================
Private Function ObterDataArquivo(ByVal nome As String, ByVal caminho As String) As Date
    Dim semExtensao As String
    Dim textoData As String

    On Error GoTo UsarModificacao
    semExtensao = Left$(nome, InStrRev(nome, ".") - 1)
    textoData = Right$(semExtensao, 10)

    If Mid$(textoData, 5, 1) = "-" And Mid$(textoData, 8, 1) = "-" Then
        ObterDataArquivo = DateSerial( _
            CInt(Left$(textoData, 4)), _
            CInt(Mid$(textoData, 6, 2)), _
            CInt(Right$(textoData, 2)))
        Exit Function
    End If

UsarModificacao:
    ObterDataArquivo = FileDateTime(caminho)
End Function

'==============================================================================
Private Function NormalizarOP(ByVal valor As Variant) As String
    Dim origem As String
    Dim digitos As String
    Dim i As Long
    Dim caractere As String

    If IsError(valor) Or IsEmpty(valor) Then Exit Function

    If IsNumeric(valor) Then
        origem = Format$(CDbl(valor), "0")
    Else
        origem = CStr(valor)
    End If

    For i = 1 To Len(origem)
        caractere = Mid$(origem, i, 1)
        If caractere >= "0" And caractere <= "9" Then
            digitos = digitos & caractere
        End If
    Next i

    If Len(digitos) = 0 Then Exit Function

    ' ATENCAO: regra de negocio - CONFIRA. Para O.P. com mais de 6 digitos o
    ' codigo original mantem "0" + os 5 PRIMEIROS digitos, o que TRUNCA a O.P.
    ' Se as suas O.P.s podem ter 7+ digitos reais, isso causa erro de
    ' correspondencia. Mantido como estava; ajuste se necessario.
    If Len(digitos) <= 6 Then
        NormalizarOP = Right$("000000" & digitos, 6)
    Else
        NormalizarOP = "0" & Left$(digitos, 5)
    End If
End Function

'==============================================================================
Private Sub GravarBase(ByVal abertas As Object, ByVal fechadas As Object)
    Dim todos As Object
    Dim chave As Variant
    Dim regA As Variant
    Dim regF As Variant
    Dim saida() As Variant
    Dim linha As Long
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim estavaProtegida As Boolean

    Set todos = CreateObject("Scripting.Dictionary")
    todos.CompareMode = vbTextCompare

    For Each chave In abertas.Keys
        todos(chave) = True
    Next chave
    For Each chave In fechadas.Keys
        todos(chave) = True
    Next chave

    If todos.Count = 0 Then
        Err.Raise vbObjectError + 1003, , "Os arquivos não possuem O.P.s válidas."
    End If

    ReDim saida(1 To todos.Count + 1, 1 To 7)
    saida(1, 1) = "OP"
    saida(1, 2) = "Emissão"
    saida(1, 3) = "Encerramento"
    saida(1, 4) = "Situação"
    saida(1, 5) = "Arquivo Emissão"
    saida(1, 6) = "Arquivo Encerramento"
    saida(1, 7) = "Atualizado em"

    linha = 2
    For Each chave In todos.Keys
        saida(linha, 1) = CStr(chave)

        If abertas.Exists(chave) Then
            regA = abertas(chave)
            saida(linha, 2) = regA(0)
            saida(linha, 5) = regA(2)
        End If

        If fechadas.Exists(chave) Then
            regF = fechadas(chave)
            If Len(CStr(saida(linha, 2))) = 0 Then
                saida(linha, 2) = regF(0)
                saida(linha, 5) = regF(3)
            End If
            saida(linha, 3) = regF(1)
            saida(linha, 6) = regF(3)
        End If

        If Len(CStr(saida(linha, 3))) > 0 Then
            saida(linha, 4) = "FECHADA"
        Else
            saida(linha, 4) = "ABERTA"
        End If

        saida(linha, 7) = Now
        linha = linha + 1
    Next chave

    If Not PlanilhaExiste(NOME_BASE) Then
        Err.Raise vbObjectError + 1004, , "A aba '" & NOME_BASE & "' não existe neste arquivo."
    End If

    Set ws = ThisWorkbook.Worksheets(NOME_BASE)
    ws.Visible = xlSheetVisible
    estavaProtegida = Desproteger(ws)   ' <<< evita erro 1004 se a aba estiver protegida

    On Error Resume Next
    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo
    On Error GoTo 0

    ws.Range("A:G").ClearContents
    ws.Range("A1").Resize(UBound(saida, 1), 7).Value = saida
    ws.Columns("A").NumberFormat = "@"
    ws.Columns("B:C").NumberFormat = "dd/mm/yyyy"
    ws.Columns("G").NumberFormat = "dd/mm/yyyy hh:mm"

    If estavaProtegida Then Reproteger ws
    ws.Visible = xlSheetVeryHidden
End Sub

'==============================================================================
Public Sub AtualizarStatusFinalizado()
    Dim nomes As Variant
    Dim limites As Variant
    Dim i As Long, linha As Long
    Dim ws As Worksheet
    Dim raw As Variant
    Dim ehData As Boolean
    Dim estavaProtegida As Boolean

    nomes = Array( _
        "Formulacao", "Enchimento-Forma", "Cord", "Go", _
        "Indumak-Minipack", "Manual", "Montagem Valo", _
        "Nacionalizacao", "Universal", "Embalagem", "Apoio")
    limites = Array(75, 75, 75, 75, 81, 75, 75, 75, 75, 75, 75)

    For i = LBound(nomes) To UBound(nomes)
        If Not PlanilhaExiste(CStr(nomes(i))) Then GoTo ProximaAba   ' evita erro 9

        Set ws = ThisWorkbook.Worksheets(CStr(nomes(i)))
        estavaProtegida = Desproteger(ws)   ' <<< evita erro 1004 ao escrever

        For linha = 4 To CLng(limites(i))
            raw = ws.Cells(linha, 11).Value2
            If Not IsError(raw) Then                       ' evita erro 13 em celulas com #N/A etc.
                If Len(CStr(raw)) > 0 Then
                    ehData = False
                    If IsNumeric(raw) Then
                        If CDbl(raw) > 0 Then ehData = True
                    ElseIf IsDate(ws.Cells(linha, 11).Value) Then
                        ehData = True
                    End If
                    If ehData Then
                        If UCase$(Trim$(CStr(ws.Cells(linha, 8).Value2))) <> "FINALIZADO" Then
                            ws.Cells(linha, 8).Value = "FINALIZADO"
                        End If
                    End If
                End If
            End If
        Next linha

        If estavaProtegida Then Reproteger ws
ProximaAba:
    Next i
End Sub

'==============================================================================
Private Sub ConfigurarListaStatus()
    Dim wsLista As Worksheet
    Dim ws As Worksheet
    Dim nomes As Variant
    Dim limites As Variant
    Dim opcoes As Variant
    Dim i As Long
    Dim rng As Range
    Dim estavaProtegida As Boolean

    If Not PlanilhaExiste("LISTAS") Then
        Err.Raise vbObjectError + 1005, , "A aba 'LISTAS' não existe neste arquivo."
    End If

    Set wsLista = ThisWorkbook.Worksheets("LISTAS")
    estavaProtegida = Desproteger(wsLista)

    opcoes = Array( _
        "ABRIR ORDEM", _
        "SEPARAÇÃO", _
        "AGUARDANDO PRODUÇÃO", _
        "PRODUZINDO", _
        "AGUARDANDO FORMULAÇÃO", _
        "AGUARDANDO ENCHIMENTO", _
        "AGUARDANDO ENDEREÇAMENTO", _
        "AGUARDANDO IMPORTAÇÃO")

    For i = LBound(opcoes) To UBound(opcoes)
        wsLista.Cells(i + 2, 1).Value = opcoes(i)
    Next i
    wsLista.Range("A10:A30").ClearContents
    If estavaProtegida Then Reproteger wsLista

    On Error Resume Next
    ThisWorkbook.Names("ListaStatusManual").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="ListaStatusManual", RefersTo:="=LISTAS!$A$2:$A$9"

    nomes = Array( _
        "Formulacao", "Enchimento-Forma", "Cord", "Go", _
        "Indumak-Minipack", "Manual", "Montagem Valo", _
        "Nacionalizacao", "Universal", "Embalagem", "Apoio")
    limites = Array(75, 75, 75, 75, 81, 75, 75, 75, 75, 75, 75)

    For i = LBound(nomes) To UBound(nomes)
        If Not PlanilhaExiste(CStr(nomes(i))) Then GoTo ProximaAba   ' evita erro 9

        Set ws = ThisWorkbook.Worksheets(CStr(nomes(i)))
        estavaProtegida = Desproteger(ws)   ' <<< Validation.Add em aba protegida = erro 1004

        Set rng = ws.Range("H4:H" & CStr(limites(i)))
        On Error Resume Next
        rng.Validation.Delete
        On Error GoTo 0
        rng.Validation.Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="=ListaStatusManual"
        rng.Validation.IgnoreBlank = True
        rng.Validation.InCellDropdown = True
        rng.Validation.ShowError = True
        rng.Validation.ErrorTitle = "Status inválido"
        rng.Validation.ErrorMessage = "Escolha um Status da lista. FINALIZADO é preenchido automaticamente."

        If estavaProtegida Then Reproteger ws
ProximaAba:
    Next i
End Sub

'==============================================================================
Private Sub RemoverConsultasIncompletas()
    On Error Resume Next
    ThisWorkbook.Connections("Query - qBASE_OP").Delete
    ThisWorkbook.Connections("Query - qOP_Fechadas").Delete
    ThisWorkbook.Connections("Query - qOP_Abertas").Delete
    ThisWorkbook.Queries("qBASE_OP").Delete
    ThisWorkbook.Queries("qOP_Fechadas").Delete
    ThisWorkbook.Queries("qOP_Abertas").Delete
    On Error GoTo 0
End Sub

'==============================================================================
'  HELPERS ACRESCENTADOS
'==============================================================================
Private Function PlanilhaExiste(ByVal nome As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    PlanilhaExiste = Not ws Is Nothing
End Function

Private Function PastaExiste(ByVal caminho As String) As Boolean
    On Error Resume Next
    PastaExiste = (Len(Dir$(caminho, vbDirectory)) > 0)
    On Error GoTo 0
End Function

' Desprotege a aba se estiver protegida. Retorna True se ela ESTAVA protegida
' (para sabermos se devemos reproteger no final).
Private Function Desproteger(ByVal ws As Worksheet) As Boolean
    Desproteger = ws.ProtectContents
    If Desproteger Then
        On Error Resume Next
        ws.Unprotect Password:=SENHA_PROTECAO
        On Error GoTo 0
    End If
End Function

Private Sub Reproteger(ByVal ws As Worksheet)
    On Error Resume Next
    ' UserInterfaceOnly:=True mantem a aba protegida para o usuario, mas
    ' permite que o codigo VBA escreva sem gerar erro 1004.
    ws.Protect Password:=SENHA_PROTECAO, DrawingObjects:=True, _
               Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub RestaurarApp()
    On Error Resume Next
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub
