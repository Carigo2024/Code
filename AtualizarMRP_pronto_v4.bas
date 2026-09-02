Attribute VB_Name = "modAtualizarMRP"
Option Explicit

' ============================================================================
' ATUALIZACAO MRP - PROTHEUS
'
' Arquivo mantido somente com caracteres ASCII para evitar corrupcao de
' literais durante a importacao pelo editor VBA. Textos exibidos com acentos
' sao montados em tempo de execucao com ChrW.
'
' Regras implementadas:
'   - 12 periodos mensais consecutivos.
'   - Uma unica linha por SKU, com bloqueio de duplicidades.
'   - Saida total = Saidas + Saida Estrutura.
'   - Somente Tipo PA gera producao; demais tipos geram compra.
'   - Cobertura = (Estoque Atual + Estoque Seguranca) / consumo diario.
'   - Quantidade sugerida = maior entre necessidade media, lote minimo e
'     lote economico. O lote economico e tratado como piso, nao como multiplo.
'   - Valor recomendado = quantidade sugerida x valor unitario, arredondado
'     em centavos antes da gravacao e conferido por SKU no destino.
'   - Base_Original, Base_Tratada e Analise_MRP sao atualizadas pela mesma carga.
' ============================================================================

Private Const ABA_FONTE As String = "Resultados"
Private Const ABA_ORIGINAL As String = "Base_Original"
Private Const ABA_TRATADA As String = "Base_Tratada"
Private Const ABA_ANALISE As String = "Analise_MRP"
Private Const ABA_RESUMO As String = "Resumo_Executivo"
Private Const ABA_PARAMETROS As String = "Parametros"
Private Const ABA_QA As String = "QA_Validacao"

Private Const NPER As Long = 12
Private Const COLUNAS_FONTE As Long = 14
Private Const COLUNAS_TRATADA As Long = 40
Private Const COLUNAS_ANALISE As Long = 20
Private Const COLUNAS_ORDENACAO As Long = 24
Private Const FATOR_EXCESSO_ANUAL As Double = 2#
Private Const PASTA_PADRAO As String = "C:\MRP\"
Private Const TOLERANCIA As Double = 0.000001
Private Const TOLERANCIA_MOEDA As Double = 0.01

Private Type TSnapshot
    Conteudo As Variant
    PrimeiraLinha As Long
    PrimeiraColuna As Long
    NumeroLinhas As Long
    NumeroColunas As Long
    Valido As Boolean
End Type

Private mEtapa As String
Private mContexto As String

' Unico ponto publico. Reassocie o botao da planilha a esta macro.
Public Sub AtualizarMRP()
    Dim caminhoArquivo As String
    Dim wbFonte As Workbook
    Dim wsFonte As Worksheet
    Dim wsOriginal As Worksheet
    Dim wsTratada As Worksheet
    Dim wsAnalise As Worksheet
    Dim wsParametros As Worksheet
    Dim wsQA As Worksheet

    Dim periodos(1 To NPER) As Date
    Dim matrizFonte As Variant
    Dim matrizTratada As Variant
    Dim matrizAnalise As Variant
    Dim ultimaLinhaFonte As Long
    Dim quantidadeProdutos As Long

    Dim snapOriginal As TSnapshot
    Dim snapTratada As TSnapshot
    Dim snapAnalise As TSnapshot
    Dim snapParametros As TSnapshot
    Dim snapQA As TSnapshot

    Dim calculoAnterior As XlCalculation
    Dim telaAnterior As Boolean
    Dim alertasAnteriores As Boolean
    Dim eventosAnteriores As Boolean
    Dim barraAnterior As Variant
    Dim segurancaAnterior As Long
    Dim segurancaAlterada As Boolean
    Dim estadoCapturado As Boolean

    Dim gravacaoIniciada As Boolean
    Dim sucesso As Boolean
    Dim cancelado As Boolean
    Dim numeroErro As Long
    Dim descricaoErro As String
    Dim origemErro As String
    Dim inicio As Double
    Dim duracao As Double

    On Error GoTo TrataErro

    mEtapa = "Inicializacao"
    mContexto = vbNullString
    inicio = Timer
    calculoAnterior = Application.Calculation
    telaAnterior = Application.ScreenUpdating
    alertasAnteriores = Application.DisplayAlerts
    eventosAnteriores = Application.EnableEvents
    barraAnterior = Application.StatusBar
    segurancaAnterior = Application.AutomationSecurity
    estadoCapturado = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "MRP: selecione o relatorio de origem."

    caminhoArquivo = SelecionarArquivo()
    If Len(caminhoArquivo) = 0 Then
        cancelado = True
        GoTo Finalizar
    End If

    If StrComp(caminhoArquivo, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1000, "AtualizarMRP", _
                  "O arquivo de origem nao pode ser a propria planilha MRP."
    End If

    Application.StatusBar = "MRP: abrindo e validando o relatorio."
    Application.AutomationSecurity = 3
    segurancaAlterada = True

    Set wbFonte = Workbooks.Open( _
        Filename:=caminhoArquivo, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        AddToMru:=False, _
        Notify:=False)

    Application.AutomationSecurity = segurancaAnterior
    segurancaAlterada = False

    Set wsFonte = ObterPlanilhaObrigatoria(wbFonte, ABA_FONTE)
    mEtapa = "Ler periodos"
    mContexto = caminhoArquivo
    LerPeriodos wsFonte, periodos

    mEtapa = "Ler matriz de origem"
    matrizFonte = LerMatrizFonte(wsFonte, ultimaLinhaFonte)

    Application.StatusBar = "MRP: processando produtos e movimentos."
    mEtapa = "Processar produtos e movimentos"
    ProcessarBlocos wsFonte, matrizFonte, periodos, _
                    matrizTratada, matrizAnalise, quantidadeProdutos

    mEtapa = "Validar matrizes em memoria"
    ValidarMatrizes matrizTratada, matrizAnalise, quantidadeProdutos

    wbFonte.Close SaveChanges:=False
    Set wbFonte = Nothing

    Set wsOriginal = ObterPlanilhaObrigatoria(ThisWorkbook, ABA_ORIGINAL)
    Set wsTratada = ObterPlanilhaObrigatoria(ThisWorkbook, ABA_TRATADA)
    Set wsAnalise = ObterPlanilhaObrigatoria(ThisWorkbook, ABA_ANALISE)
    Set wsParametros = ObterPlanilhaObrigatoria(ThisWorkbook, ABA_PARAMETROS)
    Set wsQA = ObterPlanilhaOpcional(ThisWorkbook, ABA_QA)

    mEtapa = "Capturar copia de seguranca"
    mContexto = vbNullString
    CapturarSnapshot wsOriginal, snapOriginal
    CapturarSnapshot wsTratada, snapTratada
    CapturarSnapshot wsAnalise, snapAnalise
    CapturarSnapshot wsParametros, snapParametros
    If Not wsQA Is Nothing Then CapturarSnapshot wsQA, snapQA

    Application.StatusBar = "MRP: gravando a carga validada."
    gravacaoIniciada = True

    mEtapa = "Gravar Base_Original"
    mContexto = ABA_ORIGINAL
    EscreverBaseOriginal wsOriginal, matrizFonte, ultimaLinhaFonte

    mEtapa = "Gravar Base_Tratada"
    mContexto = ABA_TRATADA
    EscreverBaseTratada wsTratada, matrizTratada, quantidadeProdutos, periodos

    mEtapa = "Gravar Analise_MRP"
    mContexto = ABA_ANALISE
    EscreverAnalise wsAnalise, matrizAnalise, quantidadeProdutos

    mEtapa = "Atualizar Parametros"
    mContexto = ABA_PARAMETROS
    AtualizarParametros wsParametros, periodos

    mEtapa = "Atualizar QA"
    mContexto = ABA_QA
    AtualizarQA wsQA, quantidadeProdutos

    mEtapa = "Calcular Resumo_Executivo"
    mContexto = ABA_RESUMO
    ThisWorkbook.Worksheets(ABA_RESUMO).Calculate

    If Not wsQA Is Nothing Then
        mEtapa = "Calcular QA_Validacao"
        mContexto = ABA_QA
        wsQA.Calculate
    End If

    Application.StatusBar = "MRP: executando reconciliacao final."
    mEtapa = "Validar destino"
    mContexto = vbNullString
    ValidarDestino wsOriginal, wsTratada, wsAnalise, wsQA, _
                   quantidadeProdutos, ultimaLinhaFonte

    sucesso = True
    gravacaoIniciada = False

Finalizar:
    On Error Resume Next

    If Not wbFonte Is Nothing Then wbFonte.Close SaveChanges:=False

    If segurancaAlterada Then
        Application.AutomationSecurity = segurancaAnterior
    End If

    If estadoCapturado Then
        Application.StatusBar = barraAnterior
        Application.Calculation = calculoAnterior
        Application.EnableEvents = eventosAnteriores
        Application.DisplayAlerts = alertasAnteriores
        Application.ScreenUpdating = telaAnterior
    End If

    Application.CutCopyMode = False
    On Error GoTo 0

    If cancelado Then
        MsgBox "Atualizacao cancelada.", vbInformation
    ElseIf sucesso Then
        duracao = Timer - inicio
        If duracao < 0 Then duracao = duracao + 86400#

        MsgBox "Analise MRP atualizada e validada." & vbCrLf & _
               CStr(quantidadeProdutos) & " produtos unicos em " & _
               Format$(duracao, "0.0") & " s.", vbInformation
    Else
        MsgBox "A atualizacao foi interrompida e os dados anteriores foram restaurados." & _
               vbCrLf & vbCrLf & _
               "Erro " & CStr(numeroErro) & ": " & descricaoErro & _
               IIf(Len(origemErro) > 0, vbCrLf & "Origem: " & origemErro, vbNullString), _
               vbCritical
    End If

    Exit Sub

TrataErro:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    origemErro = Err.Source

    If Len(mEtapa) > 0 Then
        descricaoErro = descricaoErro & vbCrLf & "Etapa: " & mEtapa
    End If
    If Len(mContexto) > 0 Then
        descricaoErro = descricaoErro & vbCrLf & "Contexto: " & mContexto
    End If

    If gravacaoIniciada Then
        On Error Resume Next
        Application.StatusBar = "MRP: restaurando a carga anterior."
        RestaurarSnapshot wsOriginal, snapOriginal
        RestaurarSnapshot wsTratada, snapTratada
        RestaurarSnapshot wsAnalise, snapAnalise
        RestaurarSnapshot wsParametros, snapParametros
        If Not wsQA Is Nothing Then RestaurarSnapshot wsQA, snapQA
        ThisWorkbook.Worksheets(ABA_RESUMO).Calculate
        On Error GoTo 0
    End If

    GoTo Finalizar
End Sub

Private Function SelecionarArquivo() As String
    Dim fd As FileDialog
    Dim pastaExiste As Boolean
    Dim titulo As String
    Dim filtro As String
    Dim escolhido As Variant

    titulo = "Selecione o novo relat" & ChrW(243) & "rio MRP do Protheus"
    filtro = "Arquivos Excel (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls"

    ' Application.FileDialog pode retornar Nothing em algumas configuracoes do
    ' Excel (add-ins, automacao, sessao sem janela ativa, politicas corporativas).
    ' Sem protecao, o bloco With gerava Erro 91. Capturamos e, se necessario,
    ' usamos Application.GetOpenFilename como alternativa.
    On Error Resume Next
    Set fd = Application.FileDialog(3)
    pastaExiste = (Len(Dir$(PASTA_PADRAO, vbDirectory)) > 0)
    On Error GoTo 0

    If Not fd Is Nothing Then
        With fd
            .Title = titulo
            .Filters.Clear
            .Filters.Add "Arquivos Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls"
            .AllowMultiSelect = False

            If pastaExiste Then .InitialFileName = PASTA_PADRAO

            If .Show = -1 Then
                SelecionarArquivo = CStr(.SelectedItems(1))
            Else
                SelecionarArquivo = vbNullString
            End If
        End With
        Exit Function
    End If

    ' Alternativa robusta (nao depende do objeto FileDialog)
    If pastaExiste Then
        On Error Resume Next
        ChDrive PASTA_PADRAO
        ChDir PASTA_PADRAO
        On Error GoTo 0
    End If
    escolhido = Application.GetOpenFilename(filtro, 1, titulo, , False)

    If VarType(escolhido) = vbBoolean Then
        SelecionarArquivo = vbNullString
    Else
        SelecionarArquivo = CStr(escolhido)
    End If
End Function

Private Function ObterPlanilhaObrigatoria( _
    ByVal wb As Workbook, _
    ByVal nomePlanilha As String) As Worksheet

    On Error Resume Next
    Set ObterPlanilhaObrigatoria = wb.Worksheets(nomePlanilha)
    On Error GoTo 0

    If ObterPlanilhaObrigatoria Is Nothing Then
        Err.Raise vbObjectError + 1001, "ObterPlanilhaObrigatoria", _
                  "A planilha obrigatoria '" & nomePlanilha & "' nao foi encontrada em '" & _
                  wb.Name & "'."
    End If
End Function

Private Function ObterPlanilhaOpcional( _
    ByVal wb As Workbook, _
    ByVal nomePlanilha As String) As Worksheet

    On Error Resume Next
    Set ObterPlanilhaOpcional = wb.Worksheets(nomePlanilha)
    On Error GoTo 0
End Function

Private Function LerMatrizFonte( _
    ByVal ws As Worksheet, _
    ByRef ultimaLinha As Long) As Variant

    Dim ultimaCelula As Range

    Set ultimaCelula = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If ultimaCelula Is Nothing Then
        Err.Raise vbObjectError + 1002, "LerMatrizFonte", _
                  "O relatorio de origem esta vazio."
    End If

    ultimaLinha = ultimaCelula.Row

    If ultimaLinha < 7 Then
        Err.Raise vbObjectError + 1003, "LerMatrizFonte", _
                  "O relatorio nao possui linhas suficientes para um bloco de produto."
    End If

    LerMatrizFonte = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(ultimaLinha, COLUNAS_FONTE)).Value2
End Function

Private Sub LerPeriodos(ByVal ws As Worksheet, ByRef periodos() As Date)
    Dim linhaPeriodo As Long
    Dim candidatoMDY(1 To NPER) As Date
    Dim candidatoDMY(1 To NPER) As Date
    Dim validoMDY As Boolean
    Dim validoDMY As Boolean
    Dim dataMDYValida As Boolean
    Dim dataDMYValida As Boolean
    Dim texto As String
    Dim c As Long
    Dim notaMDY As Long
    Dim notaDMY As Long

    linhaPeriodo = LocalizarLinhaPeriodo(ws)
    validoMDY = True
    validoDMY = True

    For c = 1 To NPER
        texto = Trim$(CStr(ws.Cells(linhaPeriodo, c + 1).Text))

        If Len(texto) = 0 Or InStr(1, texto, "#", vbBinaryCompare) > 0 Then
            If IsDate(ws.Cells(linhaPeriodo, c + 1).Value) Then
                texto = Format$(CDate(ws.Cells(linhaPeriodo, c + 1).Value), "mm/dd/yyyy")
            Else
                Err.Raise vbObjectError + 1004, "LerPeriodos", _
                          "Periodo invalido na coluna " & CStr(c + 1) & "."
            End If
        End If

        InterpretarDataAmbigua texto, _
                                candidatoMDY(c), dataMDYValida, _
                                candidatoDMY(c), dataDMYValida

        If Not dataMDYValida Then validoMDY = False
        If Not dataDMYValida Then validoDMY = False
    Next c

    If validoMDY Then notaMDY = PontuarPeriodos(candidatoMDY)
    If validoDMY Then notaDMY = PontuarPeriodos(candidatoDMY)

    If validoMDY And (Not validoDMY Or notaMDY > notaDMY) Then
        For c = 1 To NPER
            periodos(c) = candidatoMDY(c)
        Next c
    ElseIf validoDMY And (Not validoMDY Or notaDMY > notaMDY) Then
        For c = 1 To NPER
            periodos(c) = candidatoDMY(c)
        Next c
    Else
        Err.Raise vbObjectError + 1005, "LerPeriodos", _
                  "Nao foi possivel interpretar os 12 periodos como meses consecutivos."
    End If

    For c = 1 To NPER
        If Day(periodos(c)) <> 1 Then
            Err.Raise vbObjectError + 1006, "LerPeriodos", _
                      "Todos os periodos devem representar o primeiro dia do mes."
        End If

        If c > 1 Then
            If DateDiff("m", periodos(c - 1), periodos(c)) <> 1 Then
                Err.Raise vbObjectError + 1007, "LerPeriodos", _
                          "Os periodos nao sao mensais, consecutivos e crescentes."
            End If
        End If
    Next c
End Sub

Private Function LocalizarLinhaPeriodo(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim chave As String

    For r = 1 To 20
        chave = NormalizarRotulo(ws.Cells(r, 1).Value2)
        If chave = "PERIODO" Then
            LocalizarLinhaPeriodo = r
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 1008, "LocalizarLinhaPeriodo", _
              "A linha identificada por 'Periodo' nao foi encontrada nas primeiras 20 linhas."
End Function

Private Sub InterpretarDataAmbigua( _
    ByVal texto As String, _
    ByRef dataMDY As Date, _
    ByRef validoMDY As Boolean, _
    ByRef dataDMY As Date, _
    ByRef validoDMY As Boolean)

    Dim partes() As String
    Dim dataLimpa As String
    Dim parteData As String
    Dim p1 As Long
    Dim p2 As Long
    Dim ano As Long

    dataLimpa = Trim$(texto)
    dataLimpa = Replace(dataLimpa, ".", "/")
    dataLimpa = Replace(dataLimpa, "-", "/")

    If InStr(1, dataLimpa, " ", vbBinaryCompare) > 0 Then
        parteData = Split(dataLimpa, " ")(0)
    Else
        parteData = dataLimpa
    End If

    partes = Split(parteData, "/")
    If UBound(partes) - LBound(partes) <> 2 Then Exit Sub

    If Not IsNumeric(partes(0)) Then Exit Sub
    If Not IsNumeric(partes(1)) Then Exit Sub
    If Not IsNumeric(partes(2)) Then Exit Sub

    p1 = CLng(partes(0))
    p2 = CLng(partes(1))
    ano = CLng(partes(2))
    If ano < 100 Then ano = 2000 + ano

    dataMDY = CriarDataSegura(ano, p1, p2, validoMDY)
    dataDMY = CriarDataSegura(ano, p2, p1, validoDMY)
End Sub

Private Function CriarDataSegura( _
    ByVal ano As Long, _
    ByVal mes As Long, _
    ByVal dia As Long, _
    ByRef valido As Boolean) As Date

    Dim d As Date

    valido = False
    If ano < 1900 Or ano > 9999 Then Exit Function
    If mes < 1 Or mes > 12 Then Exit Function
    If dia < 1 Or dia > 31 Then Exit Function

    On Error GoTo DataInvalida
    d = DateSerial(ano, mes, dia)

    If Year(d) <> ano Then GoTo DataInvalida
    If Month(d) <> mes Then GoTo DataInvalida
    If Day(d) <> dia Then GoTo DataInvalida

    valido = True
    CriarDataSegura = d
    Exit Function

DataInvalida:
    valido = False
End Function

Private Function PontuarPeriodos(ByRef datas() As Date) As Long
    Dim i As Long
    Dim nota As Long

    For i = 1 To NPER
        If Day(datas(i)) = 1 Then
            nota = nota + 2
        Else
            nota = nota - 5
        End If

        If i > 1 Then
            If DateDiff("m", datas(i - 1), datas(i)) = 1 Then
                nota = nota + 10
            Else
                nota = nota - 20
            End If
        End If
    Next i

    PontuarPeriodos = nota
End Function

Private Sub ProcessarBlocos( _
    ByVal wsFonte As Worksheet, _
    ByRef fonte As Variant, _
    ByRef periodos() As Date, _
    ByRef tratada As Variant, _
    ByRef analise As Variant, _
    ByRef nProd As Long)

    Dim ultimaLinha As Long
    Dim r As Long
    Dim rr As Long
    Dim linhaMestre As Long
    Dim idx As Long
    Dim k As Long
    Dim quantidadeMarcadores As Long

    Dim vistos As Object
    Dim chave As String
    Dim codigo As String
    Dim descricao As String
    Dim tipo As String
    Dim natureza As String
    Dim armazem As String
    Dim um As String

    Dim leadTime As Double
    Dim loteMinimo As Double
    Dim loteEconomico As Double
    Dim estoqueSeguranca As Double
    Dim pontoPedido As Double
    Dim valorUnitario As Double
    Dim estoqueAtual As Double
    Dim recebimentos As Double
    Dim saidaTotal As Double
    Dim necessidadeTotal As Double
    Dim saldoMinimo As Double
    Dim saidaMedia As Double
    Dim necessidadeMedia As Double
    Dim cobertura As Double
    Dim quantidadeSugerida As Double
    Dim quantidadeComprar As Double
    Dim quantidadeProduzir As Double
    Dim valorRecomendado As Double

    Dim entradas(1 To NPER) As Double
    Dim saidas(1 To NPER) As Double
    Dim saidasEstrutura(1 To NPER) As Double
    Dim necessidades(1 To NPER) As Double
    Dim saldos(1 To NPER) As Double

    Dim encontrouEstoque As Boolean
    Dim encontrouEntradas As Boolean
    Dim encontrouSaidas As Boolean
    Dim encontrouEstrutura As Boolean
    Dim encontrouSaldo As Boolean
    Dim encontrouNecessidade As Boolean
    Dim ehProduzido As Boolean
    Dim dadosInvalidos As Boolean
    Dim urgente As Boolean
    Dim rotulo As String
    Dim motivo As String
    Dim status As String
    Dim prioridade As String
    Dim observacao As String
    Dim coberturaSaida As Variant
    Dim dataCritica As Variant
    Dim indiceCritico As Long
    Dim indiceNecessidade As Long

    ultimaLinha = UBound(fonte, 1)

    For r = 1 To ultimaLinha - 1
        If NormalizarRotulo(fonte(r, 1)) = "PRODUTO" Then
            If Len(LimparSKU(fonte(r + 1, 1))) > 0 Then
                quantidadeMarcadores = quantidadeMarcadores + 1
            End If
        End If
    Next r

    If quantidadeMarcadores = 0 Then
        Err.Raise vbObjectError + 1010, "ProcessarBlocos", _
                  "Nenhum bloco de produto foi encontrado."
    End If

    ReDim tratada(1 To quantidadeMarcadores, 1 To COLUNAS_TRATADA)
    ReDim analise(1 To quantidadeMarcadores, 1 To COLUNAS_ORDENACAO)

    Set vistos = CreateObject("Scripting.Dictionary")
    vistos.CompareMode = vbTextCompare

    r = 1
    Do While r <= ultimaLinha
        If NormalizarRotulo(fonte(r, 1)) <> "PRODUTO" Then
            r = r + 1
        Else
            linhaMestre = r + 1

            If linhaMestre > ultimaLinha Then
                Err.Raise vbObjectError + 1011, "ProcessarBlocos", _
                          "Cabecalho Produto sem linha mestre na linha " & CStr(r) & "."
            End If

            codigo = LimparSKU(wsFonte.Cells(linhaMestre, 1).Text)
            If Len(codigo) = 0 Or Left$(codigo, 1) = "#" Then
                codigo = LimparSKU(fonte(linhaMestre, 1))
            End If

            If Len(codigo) = 0 Then
                Err.Raise vbObjectError + 1012, "ProcessarBlocos", _
                          "Codigo de produto vazio na linha " & CStr(linhaMestre) & "."
            End If

            chave = ChaveSKU(codigo)
            If vistos.Exists(chave) Then
                Err.Raise vbObjectError + 1013, "ProcessarBlocos", _
                          "SKU duplicado na origem: '" & codigo & "'. Primeira ocorrencia no item " & _
                          CStr(vistos(chave)) & "."
            End If

            idx = idx + 1
            vistos.Add chave, idx

            descricao = TextoSeguro(fonte(linhaMestre, 2))
            tipo = UCase$(LimparEspacos(TextoSeguro(fonte(linhaMestre, 3))))
            armazem = LimparEspacos(TextoSeguro(fonte(linhaMestre, 7)))
            um = LimparEspacos(TextoSeguro(fonte(linhaMestre, 13)))

            leadTime = ValorNumerico(fonte(linhaMestre, 5), Contexto(codigo, "Lead Time"))
            loteMinimo = ValorNumerico(fonte(linhaMestre, 9), Contexto(codigo, "Lote Minimo"))
            loteEconomico = ValorNumerico(fonte(linhaMestre, 10), Contexto(codigo, "Lote Economico"))
            estoqueSeguranca = ValorNumerico(fonte(linhaMestre, 11), Contexto(codigo, "Estoque Seguranca"))
            pontoPedido = ValorNumerico(fonte(linhaMestre, 12), Contexto(codigo, "Ponto Pedido"))
            valorUnitario = ValorNumerico(fonte(linhaMestre, 14), Contexto(codigo, "Valor Unitario"))

            estoqueAtual = 0#
            recebimentos = 0#
            saidaTotal = 0#
            necessidadeTotal = 0#
            saldoMinimo = 1E+308
            indiceCritico = 0
            indiceNecessidade = 0

            encontrouEstoque = False
            encontrouEntradas = False
            encontrouSaidas = False
            encontrouEstrutura = False
            encontrouSaldo = False
            encontrouNecessidade = False

            For k = 1 To NPER
                entradas(k) = 0#
                saidas(k) = 0#
                saidasEstrutura(k) = 0#
                necessidades(k) = 0#
                saldos(k) = 0#
            Next k

            rr = linhaMestre + 1
            Do While rr <= ultimaLinha
                If NormalizarRotulo(fonte(rr, 1)) = "PRODUTO" Then Exit Do

                rotulo = NormalizarRotulo(fonte(rr, 1))

                Select Case rotulo
                    Case "+ ESTOQUE"
                        encontrouEstoque = True
                        estoqueAtual = ValorNumerico( _
                            fonte(rr, 2), Contexto(codigo, "+ Estoque"))

                    Case "+ ENTRADAS"
                        encontrouEntradas = True
                        For k = 1 To NPER
                            entradas(k) = Abs(ValorNumerico( _
                                fonte(rr, k + 1), Contexto(codigo, "+ Entradas")))
                            recebimentos = recebimentos + entradas(k)
                        Next k

                    Case "- SAIDAS"
                        encontrouSaidas = True
                        For k = 1 To NPER
                            saidas(k) = Abs(ValorNumerico( _
                                fonte(rr, k + 1), Contexto(codigo, "- Saidas")))
                            saidaTotal = saidaTotal + saidas(k)
                        Next k

                    Case "- SAIDA ESTRUTURA"
                        encontrouEstrutura = True
                        For k = 1 To NPER
                            saidasEstrutura(k) = Abs(ValorNumerico( _
                                fonte(rr, k + 1), Contexto(codigo, "- Saida Estrutura")))
                            saidaTotal = saidaTotal + saidasEstrutura(k)
                        Next k

                    Case "SALDO FINAL"
                        encontrouSaldo = True
                        For k = 1 To NPER
                            saldos(k) = ValorNumerico( _
                                fonte(rr, k + 1), Contexto(codigo, "Saldo Final"))

                            If saldos(k) < saldoMinimo Then saldoMinimo = saldos(k)
                            If saldos(k) < 0 And indiceCritico = 0 Then indiceCritico = k
                        Next k

                    Case "NECESSIDADE"
                        encontrouNecessidade = True
                        For k = 1 To NPER
                            necessidades(k) = Abs(ValorNumerico( _
                                fonte(rr, k + 1), Contexto(codigo, "Necessidade")))

                            necessidadeTotal = necessidadeTotal + necessidades(k)
                            If necessidades(k) > 0 And indiceNecessidade = 0 Then
                                indiceNecessidade = k
                            End If
                        Next k
                End Select

                rr = rr + 1
            Loop

            ValidarEstruturaBloco codigo, r, _
                                  encontrouEstoque, encontrouEntradas, _
                                  encontrouSaidas, encontrouEstrutura, _
                                  encontrouSaldo, encontrouNecessidade

            If saldoMinimo = 1E+308 Then saldoMinimo = 0#

            saidaMedia = saidaTotal / CDbl(NPER)
            necessidadeMedia = necessidadeTotal / CDbl(NPER)

            ehProduzido = (UCase$(tipo) = "PA")
            If Len(tipo) = 0 Then
                natureza = "Verificar"
            ElseIf ehProduzido Then
                natureza = "Produzir"
            Else
                natureza = "Comprar"
            End If

            motivo = vbNullString
            If Len(tipo) = 0 Then AdicionarMotivo motivo, "Tipo de item ausente"
            If leadTime < 0 Then AdicionarMotivo motivo, "Lead Time negativo"
            If loteMinimo < 0 Then AdicionarMotivo motivo, "Lote Minimo negativo"
            If loteEconomico < 0 Then AdicionarMotivo motivo, "Lote Economico negativo"
            If estoqueSeguranca < 0 Then AdicionarMotivo motivo, "Estoque Seguranca negativo"
            If pontoPedido < 0 Then AdicionarMotivo motivo, "Ponto Pedido negativo"
            If valorUnitario < 0 Then AdicionarMotivo motivo, "Valor Unitario negativo"
            If necessidadeMedia > 0 And valorUnitario <= 0 Then
                AdicionarMotivo motivo, "Valor Unitario ausente para item com necessidade"
            End If
            dadosInvalidos = (Len(motivo) > 0)

            If saidaMedia > 0 Then
                cobertura = (estoqueAtual + estoqueSeguranca) / (saidaMedia / 30#)
                coberturaSaida = cobertura
            Else
                cobertura = 0#
                coberturaSaida = vbNullString
            End If

            If indiceCritico > 0 Then
                dataCritica = periodos(indiceCritico)
            ElseIf indiceNecessidade > 0 Then
                dataCritica = periodos(indiceNecessidade)
            Else
                dataCritica = vbNullString
            End If

            quantidadeSugerida = 0#
            quantidadeComprar = 0#
            quantidadeProduzir = 0#
            urgente = False

            If Not dadosInvalidos And necessidadeMedia > 0 Then
                quantidadeSugerida = Maior3( _
                    necessidadeMedia, loteMinimo, loteEconomico)

                If ehProduzido Then
                    quantidadeProduzir = quantidadeSugerida
                Else
                    quantidadeComprar = quantidadeSugerida
                End If

                If saidaMedia > 0 Then
                    urgente = (cobertura < leadTime)
                Else
                    urgente = (saldoMinimo < 0)
                End If
            End If

            observacao = vbNullString

            If dadosInvalidos Then
                status = StatusVerificar()
                prioridade = "Verificar"
                observacao = motivo & "."

            ElseIf necessidadeMedia > 0 Then
                If ehProduzido Then
                    If urgente Then
                        status = "Produzir urgente"
                        prioridade = "Alta"
                    Else
                        status = "Produzir"
                        prioridade = PrioridadeMedia()
                    End If
                Else
                    If urgente Then
                        status = "Comprar urgente"
                        prioridade = "Alta"
                    Else
                        status = "Comprar"
                        prioridade = PrioridadeMedia()
                    End If
                End If

                If urgente And saidaMedia > 0 Then
                    observacao = "Cobertura " & Format$(cobertura, "0.0") & _
                                 " dias abaixo do lead time de " & _
                                 Format$(leadTime, "0.0") & " dias."
                ElseIf urgente Then
                    observacao = "Saldo critico; cobertura indisponivel."
                Else
                    observacao = "Necessidade mensal identificada para planejamento."
                End If

            ElseIf saldoMinimo < 0 Then
                status = "Risco de falta"
                prioridade = "Alta"
                observacao = "Saldo projetado negativo sem necessidade calculada."

            ElseIf saidaMedia = 0 Then
                status = StatusSemAcao()
                prioridade = "Baixa"
                observacao = "Sem demanda no horizonte analisado."

            ElseIf (estoqueAtual + recebimentos) > _
                   (saidaTotal * FATOR_EXCESSO_ANUAL) Then
                status = "Excesso de estoque"
                prioridade = "Baixa"
                observacao = "Estoque e recebimentos acima de dois anos de saida."

            ElseIf recebimentos > 0 Then
                status = "Coberto por recebimento planejado"
                prioridade = "Baixa"
                observacao = "Sem necessidade liquida e com recebimento planejado."

            Else
                status = "Coberto por estoque"
                prioridade = "Baixa"
                observacao = "Sem necessidade liquida no horizonte."
            End If

            valorRecomendado = ArredondarMoeda( _
                (quantidadeComprar + quantidadeProduzir) * valorUnitario)

            PreencherLinhaTratada tratada, idx, _
                                  codigo, descricao, tipo, natureza, _
                                  armazem, um, leadTime, loteMinimo, _
                                  loteEconomico, estoqueSeguranca, pontoPedido, _
                                  valorUnitario, estoqueAtual, recebimentos, _
                                  saidaTotal, necessidadeTotal, _
                                  necessidades, saldos

            PreencherLinhaAnalise analise, idx, _
                                  codigo, descricao, tipo, natureza, um, _
                                  estoqueAtual, estoqueSeguranca, recebimentos, _
                                  saidaMedia, necessidadeMedia, saldoMinimo, _
                                  leadTime, coberturaSaida, _
                                  quantidadeComprar, quantidadeProduzir, _
                                  dataCritica, status, prioridade, _
                                  valorRecomendado, observacao

            r = rr
        End If
    Loop

    If idx <> quantidadeMarcadores Then
        Err.Raise vbObjectError + 1014, "ProcessarBlocos", _
                  "A quantidade processada nao coincide com os blocos encontrados."
    End If

    nProd = idx
End Sub

Private Sub ValidarEstruturaBloco( _
    ByVal codigo As String, _
    ByVal linhaCabecalho As Long, _
    ByVal encontrouEstoque As Boolean, _
    ByVal encontrouEntradas As Boolean, _
    ByVal encontrouSaidas As Boolean, _
    ByVal encontrouEstrutura As Boolean, _
    ByVal encontrouSaldo As Boolean, _
    ByVal encontrouNecessidade As Boolean)

    Dim faltantes As String

    If Not encontrouEstoque Then AdicionarMotivo faltantes, "+ Estoque"
    If Not encontrouEntradas Then AdicionarMotivo faltantes, "+ Entradas"
    If Not encontrouSaidas Then AdicionarMotivo faltantes, "- Saidas"
    If Not encontrouEstrutura Then AdicionarMotivo faltantes, "- Saida Estrutura"
    If Not encontrouSaldo Then AdicionarMotivo faltantes, "Saldo Final"
    If Not encontrouNecessidade Then AdicionarMotivo faltantes, "Necessidade"

    If Len(faltantes) > 0 Then
        Err.Raise vbObjectError + 1015, "ValidarEstruturaBloco", _
                  "Bloco incompleto para o SKU '" & codigo & "' na linha " & _
                  CStr(linhaCabecalho) & ". Faltando: " & faltantes & "."
    End If
End Sub

Private Sub PreencherLinhaTratada( _
    ByRef dados As Variant, _
    ByVal linha As Long, _
    ByVal codigo As String, _
    ByVal descricao As String, _
    ByVal tipo As String, _
    ByVal natureza As String, _
    ByVal armazem As String, _
    ByVal um As String, _
    ByVal leadTime As Double, _
    ByVal loteMinimo As Double, _
    ByVal loteEconomico As Double, _
    ByVal estoqueSeguranca As Double, _
    ByVal pontoPedido As Double, _
    ByVal valorUnitario As Double, _
    ByVal estoqueAtual As Double, _
    ByVal recebimentos As Double, _
    ByVal saidaTotal As Double, _
    ByVal necessidadeTotal As Double, _
    ByRef necessidades() As Double, _
    ByRef saldos() As Double)

    Dim k As Long

    dados(linha, 1) = codigo
    dados(linha, 2) = descricao
    dados(linha, 3) = tipo
    dados(linha, 4) = natureza
    dados(linha, 5) = armazem
    dados(linha, 6) = um
    dados(linha, 7) = leadTime
    dados(linha, 8) = loteMinimo
    dados(linha, 9) = loteEconomico
    dados(linha, 10) = estoqueSeguranca
    dados(linha, 11) = pontoPedido
    dados(linha, 12) = valorUnitario
    dados(linha, 13) = estoqueAtual
    dados(linha, 14) = recebimentos
    dados(linha, 15) = saidaTotal
    dados(linha, 16) = necessidadeTotal

    For k = 1 To NPER
        dados(linha, 16 + k) = necessidades(k)
        dados(linha, 28 + k) = saldos(k)
    Next k
End Sub

Private Sub PreencherLinhaAnalise( _
    ByRef dados As Variant, _
    ByVal linha As Long, _
    ByVal codigo As String, _
    ByVal descricao As String, _
    ByVal tipo As String, _
    ByVal natureza As String, _
    ByVal um As String, _
    ByVal estoqueAtual As Double, _
    ByVal estoqueSeguranca As Double, _
    ByVal recebimentos As Double, _
    ByVal saidaMedia As Double, _
    ByVal necessidadeMedia As Double, _
    ByVal saldoMinimo As Double, _
    ByVal leadTime As Double, _
    ByVal cobertura As Variant, _
    ByVal quantidadeComprar As Double, _
    ByVal quantidadeProduzir As Double, _
    ByVal dataCritica As Variant, _
    ByVal status As String, _
    ByVal prioridade As String, _
    ByVal valorRecomendado As Double, _
    ByVal observacao As String)

    dados(linha, 1) = codigo
    dados(linha, 2) = descricao
    dados(linha, 3) = tipo
    dados(linha, 4) = natureza
    dados(linha, 5) = um
    dados(linha, 6) = estoqueAtual
    dados(linha, 7) = estoqueSeguranca
    dados(linha, 8) = recebimentos
    dados(linha, 9) = saidaMedia
    dados(linha, 10) = necessidadeMedia
    dados(linha, 11) = saldoMinimo
    dados(linha, 12) = leadTime
    dados(linha, 13) = cobertura
    dados(linha, 14) = quantidadeComprar
    dados(linha, 15) = quantidadeProduzir
    dados(linha, 16) = dataCritica
    dados(linha, 17) = status
    dados(linha, 18) = prioridade
    dados(linha, 19) = valorRecomendado
    dados(linha, 20) = observacao

    dados(linha, 21) = OrdemPrioridade(prioridade)
    dados(linha, 22) = valorRecomendado

    If IsDate(dataCritica) Then
        dados(linha, 23) = CDbl(CDate(dataCritica))
    Else
        dados(linha, 23) = 2958465#
    End If

    dados(linha, 24) = ChaveSKU(codigo)
End Sub

Private Sub ValidarMatrizes( _
    ByRef tratada As Variant, _
    ByRef analise As Variant, _
    ByVal nProd As Long)

    Dim vistosBase As Object
    Dim vistosAnalise As Object
    Dim i As Long
    Dim j As Long
    Dim chave As String
    Dim chaveIteracao As Variant
    Dim tipo As String
    Dim esperado As Double

    If nProd <= 0 Then
        Err.Raise vbObjectError + 1020, "ValidarMatrizes", _
                  "A carga nao possui produtos."
    End If

    If UBound(tratada, 1) <> nProd Then
        Err.Raise vbObjectError + 1021, "ValidarMatrizes", _
                  "Quantidade invalida na matriz Base_Tratada."
    End If

    If UBound(analise, 1) <> nProd Then
        Err.Raise vbObjectError + 1022, "ValidarMatrizes", _
                  "Quantidade invalida na matriz Analise_MRP."
    End If

    Set vistosBase = CreateObject("Scripting.Dictionary")
    Set vistosAnalise = CreateObject("Scripting.Dictionary")
    vistosBase.CompareMode = vbTextCompare
    vistosAnalise.CompareMode = vbTextCompare

    For i = 1 To nProd
        chave = ChaveSKU(tratada(i, 1))
        If Len(chave) = 0 Then
            Err.Raise vbObjectError + 1023, "ValidarMatrizes", _
                      "SKU vazio na Base_Tratada em memoria."
        End If
        If vistosBase.Exists(chave) Then
            Err.Raise vbObjectError + 1024, "ValidarMatrizes", _
                      "SKU duplicado na Base_Tratada em memoria: " & CStr(tratada(i, 1))
        End If
        vistosBase.Add chave, i

        chave = ChaveSKU(analise(i, 1))
        If vistosAnalise.Exists(chave) Then
            Err.Raise vbObjectError + 1025, "ValidarMatrizes", _
                      "SKU duplicado na Analise_MRP em memoria: " & CStr(analise(i, 1))
        End If
        vistosAnalise.Add chave, i

        tipo = UCase$(LimparEspacos(CStr(analise(i, 3))))
        If tipo = "PA" Then
            If Abs(CDbl(analise(i, 14))) > TOLERANCIA Then
                Err.Raise vbObjectError + 1026, "ValidarMatrizes", _
                          "Item PA com quantidade de compra: " & CStr(analise(i, 1))
            End If
        Else
            If Abs(CDbl(analise(i, 15))) > TOLERANCIA Then
                Err.Raise vbObjectError + 1027, "ValidarMatrizes", _
                          "Item nao-PA com quantidade de producao: " & CStr(analise(i, 1))
            End If
        End If

        If Not StatusValido(CStr(analise(i, 17))) Then
            Err.Raise vbObjectError + 1028, "ValidarMatrizes", _
                      "Status invalido para o SKU " & CStr(analise(i, 1)) & "."
        End If

        If Not PrioridadeValida(CStr(analise(i, 18))) Then
            Err.Raise vbObjectError + 1029, "ValidarMatrizes", _
                      "Prioridade invalida para o SKU " & CStr(analise(i, 1)) & "."
        End If

        esperado = ArredondarMoeda( _
            (CDbl(analise(i, 14)) + CDbl(analise(i, 15))) * _
            CDbl(tratada(i, 12)))

        If Abs(CDbl(analise(i, 19)) - esperado) > TOLERANCIA_MOEDA Then
            Err.Raise vbObjectError + 1030, "ValidarMatrizes", _
                      "Valor recomendado divergente para o SKU " & _
                      CStr(analise(i, 1)) & ". Gravado=" & _
                      CStr(analise(i, 19)) & "; esperado=" & _
                      CStr(esperado) & "."
        End If

        For j = 1 To COLUNAS_TRATADA
            If IsError(tratada(i, j)) Then
                Err.Raise vbObjectError + 1031, "ValidarMatrizes", _
                          "Erro de celula na matriz tratada para o SKU " & _
                          CStr(tratada(i, 1)) & "."
            End If
        Next j

        For j = 1 To COLUNAS_ANALISE
            If IsError(analise(i, j)) Then
                Err.Raise vbObjectError + 1032, "ValidarMatrizes", _
                          "Erro de celula na matriz de analise para o SKU " & _
                          CStr(analise(i, 1)) & "."
            End If
        Next j
    Next i

    If vistosBase.Count <> vistosAnalise.Count Then
        Err.Raise vbObjectError + 1033, "ValidarMatrizes", _
                  "As matrizes nao possuem a mesma quantidade de SKUs."
    End If

    For Each chaveIteracao In vistosBase.Keys
        If Not vistosAnalise.Exists(CStr(chaveIteracao)) Then
            Err.Raise vbObjectError + 1034, "ValidarMatrizes", _
                      "SKU ausente na matriz de analise: " & CStr(chaveIteracao)
        End If
    Next chaveIteracao
End Sub

Private Sub EscreverBaseOriginal( _
    ByVal ws As Worksheet, _
    ByRef fonte As Variant, _
    ByVal ultimaLinha As Long)

    Dim ultimaUsada As Long

    ultimaUsada = UltimaLinhaUsada(ws)
    If ultimaUsada > 0 Then
        ws.Range("A1:N" & CStr(ultimaUsada)).ClearContents
    End If

    ws.Range("A1").Resize(ultimaLinha, COLUNAS_FONTE).Value2 = fonte
End Sub

Private Sub EscreverBaseTratada( _
    ByVal ws As Worksheet, _
    ByRef dados As Variant, _
    ByVal nProd As Long, _
    ByRef periodos() As Date)

    Dim cabecalho As Variant
    Dim ultimaUsada As Long

    ultimaUsada = UltimaLinhaUsada(ws)
    If ultimaUsada < 2 Then ultimaUsada = 2

    ws.Range("A2:AN" & CStr(ultimaUsada)).ClearContents

    cabecalho = CabecalhoBaseTratada(periodos)
    ws.Range("A1:AN1").Value2 = cabecalho

    ws.Range("A2:F" & CStr(nProd + 1)).NumberFormat = "@"
    ws.Range("A2").Resize(nProd, COLUNAS_TRATADA).Value2 = dados

    ws.Range("G2:K" & CStr(nProd + 1)).NumberFormat = "#,##0.0"
    ws.Range("L2:L" & CStr(nProd + 1)).NumberFormat = """R$"" #,##0.00"
    ws.Range("M2:AN" & CStr(nProd + 1)).NumberFormat = "#,##0.0"

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:AN" & CStr(nProd + 1)).AutoFilter
    On Error GoTo 0
End Sub

Private Sub EscreverAnalise( _
    ByVal ws As Worksheet, _
    ByRef dados As Variant, _
    ByVal nProd As Long)

    Dim saida() As Variant
    Dim cabecalho As Variant
    Dim ultimaUsada As Long
    Dim i As Long
    Dim j As Long

    OrdenarAnalise dados, nProd

    ultimaUsada = UltimaLinhaUsada(ws)
    If ultimaUsada < 3 Then ultimaUsada = 3

    ws.Range("A3:T" & CStr(ultimaUsada)).ClearContents

    cabecalho = CabecalhoAnalise()
    ws.Range("A2:T2").Value2 = cabecalho

    ReDim saida(1 To nProd, 1 To COLUNAS_ANALISE)
    For i = 1 To nProd
        For j = 1 To COLUNAS_ANALISE
            saida(i, j) = dados(i, j)
        Next j
    Next i

    ws.Range("A3:E" & CStr(nProd + 2)).NumberFormat = "@"
    ws.Range("Q3:R" & CStr(nProd + 2)).NumberFormat = "@"
    ws.Range("T3:T" & CStr(nProd + 2)).NumberFormat = "@"
    ws.Range("A3").Resize(nProd, COLUNAS_ANALISE).Value2 = saida

    ws.Range("F3:H" & CStr(nProd + 2)).NumberFormat = "#,##0.0"
    ws.Range("I3:O" & CStr(nProd + 2)).NumberFormat = "#,##0.0"
    ws.Range("P3:P" & CStr(nProd + 2)).NumberFormat = "dd/mm/yyyy"
    ws.Range("S3:S" & CStr(nProd + 2)).NumberFormat = """R$"" #,##0.00"

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A2:T" & CStr(nProd + 2)).AutoFilter
    On Error GoTo 0
End Sub

Private Function CabecalhoBaseTratada(ByRef periodos() As Date) As Variant
    Dim h(1 To 1, 1 To COLUNAS_TRATADA) As Variant
    Dim k As Long

    h(1, 1) = "Produto"
    h(1, 2) = "Descri" & ChrW(231) & ChrW(227) & "o"
    h(1, 3) = "Tipo"
    h(1, 4) = "Natureza"
    h(1, 5) = "Armaz" & ChrW(233) & "m"
    h(1, 6) = "UM"
    h(1, 7) = "Lead Time"
    h(1, 8) = "Lote M" & ChrW(237) & "nimo"
    h(1, 9) = "Lote Econ" & ChrW(244) & "mico"
    h(1, 10) = "Estoque Seguran" & ChrW(231) & "a"
    h(1, 11) = "Ponto Pedido"
    h(1, 12) = "Valor Unit" & ChrW(225) & "rio"
    h(1, 13) = "Estoque Atual"
    h(1, 14) = "Recebimentos Planejados"
    h(1, 15) = "Sa" & ChrW(237) & "das Totais"
    h(1, 16) = "Necessidade Total"

    For k = 1 To NPER
        h(1, 16 + k) = "Nec " & Format$(periodos(k), "mm/yy")
        h(1, 28 + k) = "Saldo " & Format$(periodos(k), "mm/yy")
    Next k

    CabecalhoBaseTratada = h
End Function

Private Function CabecalhoAnalise() As Variant
    Dim h(1 To 1, 1 To COLUNAS_ANALISE) As Variant

    h(1, 1) = "Produto"
    h(1, 2) = "Descri" & ChrW(231) & ChrW(227) & "o"
    h(1, 3) = "Tipo"
    h(1, 4) = "Natureza"
    h(1, 5) = "UM"
    h(1, 6) = "Estoque Atual"
    h(1, 7) = "Estoque Seguran" & ChrW(231) & "a"
    h(1, 8) = "Recebimentos Planej."
    h(1, 9) = "Sa" & ChrW(237) & "da M" & ChrW(233) & "dia/m" & ChrW(234) & "s"
    h(1, 10) = "Necessidade M" & ChrW(233) & "dia/m" & ChrW(234) & "s"
    h(1, 11) = "Saldo Projet. M" & ChrW(237) & "n."
    h(1, 12) = "Lead Time (dias)"
    h(1, 13) = "Cobertura (dias)"
    h(1, 14) = "Qtd Comprar/m" & ChrW(234) & "s"
    h(1, 15) = "Qtd Produzir/m" & ChrW(234) & "s"
    h(1, 16) = "Data Cr" & ChrW(237) & "tica"
    h(1, 17) = "Status"
    h(1, 18) = "Prioridade"
    h(1, 19) = "Valor Necess./m" & ChrW(234) & "s (R$)"
    h(1, 20) = "Observa" & ChrW(231) & ChrW(227) & "o da An" & _
               ChrW(225) & "lise"

    CabecalhoAnalise = h
End Function

Private Sub AtualizarParametros( _
    ByVal ws As Worksheet, _
    ByRef periodos() As Date)

    ws.Range("B3").Value = periodos(1)
    ws.Range("B3").NumberFormat = "dd/mm/yyyy"
    ws.Range("B4").Value2 = CStr(NPER) & " periodos (" & _
                            Format$(periodos(1), "dd/mm/yyyy") & " a " & _
                            Format$(periodos(NPER), "dd/mm/yyyy") & ")"
End Sub

Private Sub AtualizarQA(ByVal ws As Worksheet, ByVal nProd As Long)
    If ws Is Nothing Then Exit Sub

    ws.Range("A2").Value2 = "Carga reconstruida em " & _
                            Format$(Now, "dd/mm/yyyy hh:nn") & _
                            " a partir da Base_Original."

    ws.Range("C5").Value2 = nProd
    ws.Range("C6").Value2 = nProd
    ws.Range("C8").Value2 = nProd
    ws.Range("C12").Value2 = nProd
    ws.Range("C15").Value2 = nProd
    ws.Range("C16").Value2 = nProd
End Sub

Private Sub ValidarDestino( _
    ByVal wsOriginal As Worksheet, _
    ByVal wsTratada As Worksheet, _
    ByVal wsAnalise As Worksheet, _
    ByVal wsQA As Worksheet, _
    ByVal nProd As Long, _
    ByVal ultimaLinhaFonte As Long)

    Dim dadosBase As Variant
    Dim analise As Variant
    Dim mapaBase As Object
    Dim mapaAnalise As Object
    Dim chave As String
    Dim skuAtual As String
    Dim naturezaAnalise As String
    Dim naturezaBase As String
    Dim statusAtual As String
    Dim resultadoQA As String
    Dim i As Long
    Dim j As Long
    Dim linhaBase As Long
    Dim minimo As Double
    Dim valorAtual As Double
    Dim esperado As Double
    Dim valorObservado As Double
    Dim quantidadeComprar As Double
    Dim quantidadeProduzir As Double
    Dim precoUnitario As Double
    Dim ultima As Long
    Dim totalResumo As Long

    mEtapa = "ValidarDestino - Base_Original"
    mContexto = ABA_ORIGINAL

    If UltimaLinhaUsada(wsOriginal) < ultimaLinhaFonte Then
        Err.Raise vbObjectError + 1040, "ValidarDestino", _
                  "A Base_Original foi gravada de forma incompleta."
    End If

    dadosBase = wsTratada.Range("A2:AN" & CStr(nProd + 1)).Value2
    analise = wsAnalise.Range("A3:T" & CStr(nProd + 2)).Value2

    Set mapaBase = CreateObject("Scripting.Dictionary")
    Set mapaAnalise = CreateObject("Scripting.Dictionary")
    mapaBase.CompareMode = vbTextCompare
    mapaAnalise.CompareMode = vbTextCompare

    mEtapa = "ValidarDestino - Base_Tratada"

    For i = 1 To nProd
        mContexto = ABA_TRATADA & "!linha " & CStr(i + 1)

        For j = 1 To COLUNAS_TRATADA
            If IsError(dadosBase(i, j)) Then
                Err.Raise vbObjectError + 1042, "ValidarDestino", _
                          "Erro de celula na Base_Tratada em " & _
                          EnderecoCelula(i + 1, j) & "."
            End If
        Next j

        chave = ChaveSKU(dadosBase(i, 1))
        If Len(chave) = 0 Or mapaBase.Exists(chave) Then
            Err.Raise vbObjectError + 1041, "ValidarDestino", _
                      "SKU vazio ou duplicado na Base_Tratada."
        End If
        mapaBase.Add chave, i
    Next i

    mEtapa = "ValidarDestino - Analise_MRP"

    For i = 1 To nProd
        mContexto = ABA_ANALISE & "!linha " & CStr(i + 2)

        For j = 1 To COLUNAS_ANALISE
            If IsError(analise(i, j)) Then
                Err.Raise vbObjectError + 1048, "ValidarDestino", _
                          "Erro de celula na Analise_MRP em " & _
                          EnderecoCelula(i + 2, j) & "."
            End If
        Next j

        skuAtual = LimparSKU(analise(i, 1))
        chave = ChaveSKU(analise(i, 1))

        If Len(chave) = 0 Or mapaAnalise.Exists(chave) Then
            Err.Raise vbObjectError + 1043, "ValidarDestino", _
                      "SKU vazio ou duplicado na Analise_MRP."
        End If
        mapaAnalise.Add chave, i

        If Not mapaBase.Exists(chave) Then
            Err.Raise vbObjectError + 1044, "ValidarDestino", _
                      "SKU da Analise_MRP ausente na Base_Tratada: " & skuAtual
        End If

        linhaBase = CLng(mapaBase(chave))
        mContexto = "SKU " & skuAtual

        naturezaAnalise = TextoDestino( _
            analise(i, 4), ABA_ANALISE & "!D" & CStr(i + 2) & _
            " - Natureza - SKU " & skuAtual)
        naturezaBase = TextoDestino( _
            dadosBase(linhaBase, 4), ABA_TRATADA & "!D" & _
            CStr(linhaBase + 1) & " - Natureza - SKU " & skuAtual)

        If naturezaAnalise <> naturezaBase Then
            Err.Raise vbObjectError + 1045, "ValidarDestino", _
                      "Natureza divergente para o SKU " & skuAtual
        End If

        CompararNumero analise(i, 6), dadosBase(linhaBase, 13), _
                       "Estoque Atual", skuAtual
        CompararNumero analise(i, 7), dadosBase(linhaBase, 10), _
                       "Estoque Seguranca", skuAtual
        CompararNumero analise(i, 8), dadosBase(linhaBase, 14), _
                       "Recebimentos", skuAtual
        CompararNumero analise(i, 9), _
                       NumeroDestino(dadosBase(linhaBase, 15), _
                           ABA_TRATADA & "!O" & CStr(linhaBase + 1) & _
                           " - Saidas Totais - SKU " & skuAtual) / NPER, _
                       "Saida Media", skuAtual
        CompararNumero analise(i, 10), _
                       NumeroDestino(dadosBase(linhaBase, 16), _
                           ABA_TRATADA & "!P" & CStr(linhaBase + 1) & _
                           " - Necessidade Total - SKU " & skuAtual) / NPER, _
                       "Necessidade Media", skuAtual
        CompararNumero analise(i, 12), dadosBase(linhaBase, 7), _
                       "Lead Time", skuAtual

        minimo = NumeroDestino( _
            dadosBase(linhaBase, 29), ABA_TRATADA & "!AC" & _
            CStr(linhaBase + 1) & " - Saldo 1 - SKU " & skuAtual)

        For j = 30 To 40
            valorAtual = NumeroDestino( _
                dadosBase(linhaBase, j), ABA_TRATADA & "!" & _
                EnderecoCelula(linhaBase + 1, j) & _
                " - Saldo projetado - SKU " & skuAtual)
            If valorAtual < minimo Then minimo = valorAtual
        Next j

        CompararNumero analise(i, 11), minimo, _
                       "Saldo Minimo", skuAtual

        quantidadeComprar = NumeroDestino( _
            analise(i, 14), ABA_ANALISE & "!N" & CStr(i + 2) & _
            " - Quantidade comprar - SKU " & skuAtual)
        quantidadeProduzir = NumeroDestino( _
            analise(i, 15), ABA_ANALISE & "!O" & CStr(i + 2) & _
            " - Quantidade produzir - SKU " & skuAtual)
        precoUnitario = NumeroDestino( _
            dadosBase(linhaBase, 12), ABA_TRATADA & "!L" & _
            CStr(linhaBase + 1) & " - Valor unitario - SKU " & skuAtual)
        valorObservado = NumeroDestino( _
            analise(i, 19), ABA_ANALISE & "!S" & CStr(i + 2) & _
            " - Valor recomendado - SKU " & skuAtual)

        esperado = ArredondarMoeda( _
            (quantidadeComprar + quantidadeProduzir) * precoUnitario)

        If Abs(valorObservado - esperado) > TOLERANCIA_MOEDA Then
            mContexto = "SKU " & skuAtual & _
                        "; gravado=" & CStr(valorObservado) & _
                        "; esperado=" & CStr(esperado) & _
                        "; quantidade=" & _
                        CStr(quantidadeComprar + quantidadeProduzir) & _
                        "; preco=" & CStr(precoUnitario)

            Err.Raise vbObjectError + 1046, "ValidarDestino", _
                      "Valor recomendado divergente para o SKU " & skuAtual
        End If

        statusAtual = TextoDestino( _
            analise(i, 17), ABA_ANALISE & "!Q" & CStr(i + 2) & _
            " - Status - SKU " & skuAtual)

        If Not StatusValido(statusAtual) Then
            Err.Raise vbObjectError + 1047, "ValidarDestino", _
                      "Status invalido na Analise_MRP para o SKU " & skuAtual
        End If
    Next i

    If mapaBase.Count <> nProd Or mapaAnalise.Count <> nProd Then
        Err.Raise vbObjectError + 1049, "ValidarDestino", _
                  "Contagem final de produtos invalida."
    End If

    ultima = UltimaLinhaUsada(wsTratada)
    If ultima > nProd + 1 Then
        If Application.WorksheetFunction.CountA( _
            wsTratada.Range("A" & CStr(nProd + 2) & ":AN" & CStr(ultima))) > 0 Then
            Err.Raise vbObjectError + 1050, "ValidarDestino", _
                      "Existem residuos abaixo da Base_Tratada."
        End If
    End If

    ultima = UltimaLinhaUsada(wsAnalise)
    If ultima > nProd + 2 Then
        If Application.WorksheetFunction.CountA( _
            wsAnalise.Range("A" & CStr(nProd + 3) & ":T" & CStr(ultima))) > 0 Then
            Err.Raise vbObjectError + 1051, "ValidarDestino", _
                      "Existem residuos abaixo da Analise_MRP."
        End If
    End If

    mEtapa = "ValidarDestino - Resumo_Executivo"
    totalResumo = InteiroDestino( _
        ThisWorkbook.Worksheets(ABA_RESUMO).Range("A5").Value2, _
        ABA_RESUMO & "!A5 - Total de itens")

    If totalResumo <> nProd Then
        mContexto = ABA_RESUMO & "!A5=" & CStr(totalResumo) & _
                    "; esperado=" & CStr(nProd)
        Err.Raise vbObjectError + 1052, "ValidarDestino", _
                  "O total de itens no Resumo_Executivo nao coincide com a analise."
    End If

    If Not wsQA Is Nothing Then
        mEtapa = "ValidarDestino - QA_Validacao"
        wsQA.Calculate

        resultadoQA = TextoDestino( _
            wsQA.Range("B20").Value2, _
            ABA_QA & "!B20 - Resultado final do QA")

        If UCase$(Trim$(resultadoQA)) <> "APROVADO" Then
            mContexto = ABA_QA & "!B20='" & resultadoQA & "'"
            Err.Raise vbObjectError + 1053, "ValidarDestino", _
                      "A aba QA_Validacao nao aprovou a carga."
        End If
    End If
End Sub

Private Sub CompararNumero( _
    ByVal observado As Variant, _
    ByVal esperado As Variant, _
    ByVal campo As String, _
    ByVal codigo As String)

    Dim numeroObservado As Double
    Dim numeroEsperado As Double

    numeroObservado = NumeroDestino( _
        observado, campo & " observado - SKU " & codigo)
    numeroEsperado = NumeroDestino( _
        esperado, campo & " esperado - SKU " & codigo)

    If Abs(numeroObservado - numeroEsperado) > TOLERANCIA Then
        mContexto = "SKU " & codigo & "; campo=" & campo & _
                    "; observado=" & CStr(numeroObservado) & _
                    "; esperado=" & CStr(numeroEsperado)

        Err.Raise vbObjectError + 1061, "CompararNumero", _
                  "Divergencia no campo " & campo & _
                  " para o SKU " & codigo & "."
    End If
End Sub

Private Function NumeroDestino( _
    ByVal valor As Variant, _
    ByVal contextoCampo As String) As Double

    Dim descricaoOriginal As String

    mContexto = contextoCampo

    If IsError(valor) Then
        Err.Raise vbObjectError + 1080, "NumeroDestino", _
                  "Erro de formula onde era esperado um numero."
    End If

    If IsNull(valor) Or IsEmpty(valor) Then
        Err.Raise vbObjectError + 1081, "NumeroDestino", _
                  "Valor numerico obrigatorio ausente."
    End If

    If IsObject(valor) Or IsArray(valor) Then
        Err.Raise vbObjectError + 1082, "NumeroDestino", _
                  "Tipo incompativel com numero: " & TypeName(valor) & "."
    End If

    If VarType(valor) = vbString Then
        If Len(Trim$(CStr(valor))) = 0 Then
            Err.Raise vbObjectError + 1081, "NumeroDestino", _
                      "Valor numerico obrigatorio vazio."
        End If

        NumeroDestino = ValorNumerico(valor, contextoCampo)
        Exit Function
    End If

    If Not IsNumeric(valor) Then
        Err.Raise vbObjectError + 1083, "NumeroDestino", _
                  "Tipo incompativel com numero: " & _
                  DescreverValor(valor) & "."
    End If

    On Error GoTo FalhaConversao
    NumeroDestino = CDbl(valor)
    On Error GoTo 0
    Exit Function

FalhaConversao:
    descricaoOriginal = Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 1084, "NumeroDestino", _
              "Falha ao converter numero: " & descricaoOriginal
End Function

Private Function InteiroDestino( _
    ByVal valor As Variant, _
    ByVal contextoCampo As String) As Long

    Dim numero As Double

    numero = NumeroDestino(valor, contextoCampo)

    If numero <> Fix(numero) Or _
       numero < -2147483648# Or _
       numero > 2147483647# Then
        Err.Raise vbObjectError + 1085, "InteiroDestino", _
                  "Valor nao representa um inteiro Long valido."
    End If

    InteiroDestino = CLng(numero)
End Function

Private Function TextoDestino( _
    ByVal valor As Variant, _
    ByVal contextoCampo As String) As String

    mContexto = contextoCampo

    If IsError(valor) Then
        Err.Raise vbObjectError + 1086, "TextoDestino", _
                  "Erro de formula onde era esperado texto."
    End If

    If IsNull(valor) Or IsEmpty(valor) Then
        Err.Raise vbObjectError + 1087, "TextoDestino", _
                  "Texto obrigatorio ausente."
    End If

    If IsObject(valor) Or IsArray(valor) Then
        Err.Raise vbObjectError + 1088, "TextoDestino", _
                  "Tipo incompativel com texto: " & TypeName(valor) & "."
    End If

    TextoDestino = CStr(valor)

    If Len(Trim$(TextoDestino)) = 0 Then
        Err.Raise vbObjectError + 1087, "TextoDestino", _
                  "Texto obrigatorio vazio."
    End If
End Function

Private Function DescreverValor(ByVal valor As Variant) As String
    On Error GoTo Falha

    If IsError(valor) Then
        DescreverValor = "<erro de formula>"
    ElseIf IsNull(valor) Then
        DescreverValor = "<Null>"
    ElseIf IsEmpty(valor) Then
        DescreverValor = "<Empty>"
    ElseIf IsObject(valor) Or IsArray(valor) Then
        DescreverValor = "<" & TypeName(valor) & ">"
    Else
        DescreverValor = "'" & Left$(CStr(valor), 100) & _
                         "' (" & TypeName(valor) & ")"
    End If
    Exit Function

Falha:
    DescreverValor = "<valor nao representavel>"
End Function

Private Function EnderecoCelula( _
    ByVal linha As Long, _
    ByVal coluna As Long) As String

    EnderecoCelula = ColunaExcel(coluna) & CStr(linha)
End Function

Private Function ColunaExcel(ByVal coluna As Long) As String
    Dim resultado As String
    Dim atual As Long

    atual = coluna

    Do While atual > 0
        resultado = Chr$(((atual - 1) Mod 26) + 65) & resultado
        atual = (atual - 1) \ 26
    Loop

    ColunaExcel = resultado
End Function

Private Sub OrdenarAnalise(ByRef dados As Variant, ByVal nProd As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim trocar As Boolean
    Dim temporario As Variant

    For i = 1 To nProd - 1
        For j = 1 To nProd - i
            trocar = LinhaDeveTrocar(dados, j, j + 1)

            If trocar Then
                For k = 1 To COLUNAS_ORDENACAO
                    temporario = dados(j, k)
                    dados(j, k) = dados(j + 1, k)
                    dados(j + 1, k) = temporario
                Next k
            End If
        Next j
    Next i
End Sub

Private Function LinhaDeveTrocar( _
    ByRef dados As Variant, _
    ByVal linhaA As Long, _
    ByVal linhaB As Long) As Boolean

    If CLng(dados(linhaA, 21)) > CLng(dados(linhaB, 21)) Then
        LinhaDeveTrocar = True
        Exit Function
    End If

    If CLng(dados(linhaA, 21)) < CLng(dados(linhaB, 21)) Then Exit Function

    If CDbl(dados(linhaA, 22)) < CDbl(dados(linhaB, 22)) Then
        LinhaDeveTrocar = True
        Exit Function
    End If

    If CDbl(dados(linhaA, 22)) > CDbl(dados(linhaB, 22)) Then Exit Function

    If CDbl(dados(linhaA, 23)) > CDbl(dados(linhaB, 23)) Then
        LinhaDeveTrocar = True
        Exit Function
    End If

    If CDbl(dados(linhaA, 23)) < CDbl(dados(linhaB, 23)) Then Exit Function

    LinhaDeveTrocar = _
        (StrComp(CStr(dados(linhaA, 24)), _
                 CStr(dados(linhaB, 24)), vbTextCompare) > 0)
End Function

Private Function OrdemPrioridade(ByVal prioridade As String) As Long
    Select Case prioridade
        Case "Alta"
            OrdemPrioridade = 0
        Case PrioridadeMedia()
            OrdemPrioridade = 1
        Case "Verificar"
            OrdemPrioridade = 2
        Case "Baixa"
            OrdemPrioridade = 3
        Case Else
            OrdemPrioridade = 4
    End Select
End Function

Private Function StatusValido(ByVal status As String) As Boolean
    Select Case status
        Case StatusVerificar(), _
             "Comprar urgente", _
             "Produzir urgente", _
             "Risco de falta", _
             "Comprar", _
             "Produzir", _
             "Coberto por recebimento planejado", _
             "Coberto por estoque", _
             "Excesso de estoque", _
             StatusSemAcao()
            StatusValido = True
    End Select
End Function

Private Function PrioridadeValida(ByVal prioridade As String) As Boolean
    Select Case prioridade
        Case "Alta", PrioridadeMedia(), "Baixa", "Verificar"
            PrioridadeValida = True
    End Select
End Function

Private Function StatusVerificar() As String
    StatusVerificar = "Verificar cadastro ou par" & ChrW(226) & "metro"
End Function

Private Function StatusSemAcao() As String
    StatusSemAcao = "Sem a" & ChrW(231) & ChrW(227) & "o imediata"
End Function

Private Function PrioridadeMedia() As String
    PrioridadeMedia = "M" & ChrW(233) & "dia"
End Function

Private Function ArredondarMoeda(ByVal valor As Double) As Double
    ArredondarMoeda = Application.WorksheetFunction.Round(valor, 2)
End Function

Private Function Maior3( _
    ByVal a As Double, _
    ByVal b As Double, _
    ByVal c As Double) As Double

    Maior3 = a
    If b > Maior3 Then Maior3 = b
    If c > Maior3 Then Maior3 = c
End Function

Private Function ValorNumerico( _
    ByVal valor As Variant, _
    ByVal contextoCampo As String) As Double

    Dim texto As String
    Dim normalizado As String
    Dim separadorDecimal As String
    Dim posVirgula As Long
    Dim posPonto As Long
    Dim negativoParenteses As Boolean

    If IsError(valor) Then
        Err.Raise vbObjectError + 1070, "ValorNumerico", _
                  "Erro de celula em " & contextoCampo & "."
    End If

    If IsEmpty(valor) Or IsNull(valor) Then
        ValorNumerico = 0#
        Exit Function
    End If

    If IsNumeric(valor) Then
        ValorNumerico = CDbl(valor)
        Exit Function
    End If

    texto = Trim$(CStr(valor))
    If Len(texto) = 0 Then
        ValorNumerico = 0#
        Exit Function
    End If

    texto = Replace(texto, ChrW(160), vbNullString)
    texto = Replace(texto, " ", vbNullString)
    texto = Replace(texto, "R$", vbNullString)

    If Left$(texto, 1) = "(" And Right$(texto, 1) = ")" Then
        negativoParenteses = True
        texto = Mid$(texto, 2, Len(texto) - 2)
    End If

    If IsNumeric(texto) Then
        ValorNumerico = CDbl(texto)
        If negativoParenteses Then ValorNumerico = -Abs(ValorNumerico)
        Exit Function
    End If

    separadorDecimal = Application.International(xlDecimalSeparator)
    posVirgula = InStrRev(texto, ",")
    posPonto = InStrRev(texto, ".")
    normalizado = texto

    If posVirgula > 0 And posPonto > 0 Then
        If posVirgula > posPonto Then
            normalizado = Replace(normalizado, ".", vbNullString)
            normalizado = Replace(normalizado, ",", separadorDecimal)
        Else
            normalizado = Replace(normalizado, ",", vbNullString)
            normalizado = Replace(normalizado, ".", separadorDecimal)
        End If
    ElseIf posVirgula > 0 Then
        If separadorDecimal <> "," Then
            normalizado = Replace(normalizado, ",", separadorDecimal)
        End If
    ElseIf posPonto > 0 Then
        If separadorDecimal <> "." Then
            normalizado = Replace(normalizado, ".", separadorDecimal)
        End If
    End If

    If Not IsNumeric(normalizado) Then
        Err.Raise vbObjectError + 1071, "ValorNumerico", _
                  "Valor nao numerico em " & contextoCampo & ": '" & texto & "'."
    End If

    ValorNumerico = CDbl(normalizado)
    If negativoParenteses Then ValorNumerico = -Abs(ValorNumerico)
End Function

Private Function NormalizarRotulo(ByVal valor As Variant) As String
    Dim s As String

    If IsError(valor) Or IsEmpty(valor) Or IsNull(valor) Then Exit Function

    s = CStr(valor)
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, ChrW(8203), vbNullString)
    s = Replace(s, ChrW(8204), vbNullString)
    s = Replace(s, ChrW(8205), vbNullString)
    s = Replace(s, ChrW(65279), vbNullString)

    ' Corrige sequencias comuns de UTF-8 interpretadas como Windows-1252.
    s = Replace(s, ChrW(195) & ChrW(161), "a")
    s = Replace(s, ChrW(195) & ChrW(169), "e")
    s = Replace(s, ChrW(195) & ChrW(173), "i")
    s = Replace(s, ChrW(195) & ChrW(179), "o")
    s = Replace(s, ChrW(195) & ChrW(186), "u")
    s = Replace(s, ChrW(195) & ChrW(167), "c")
    s = Replace(s, ChrW(195) & ChrW(163), "a")

    s = UCase$(s)

    s = Replace(s, ChrW(192), "A")
    s = Replace(s, ChrW(193), "A")
    s = Replace(s, ChrW(194), "A")
    s = Replace(s, ChrW(195), "A")
    s = Replace(s, ChrW(196), "A")
    s = Replace(s, ChrW(199), "C")
    s = Replace(s, ChrW(200), "E")
    s = Replace(s, ChrW(201), "E")
    s = Replace(s, ChrW(202), "E")
    s = Replace(s, ChrW(203), "E")
    s = Replace(s, ChrW(204), "I")
    s = Replace(s, ChrW(205), "I")
    s = Replace(s, ChrW(206), "I")
    s = Replace(s, ChrW(207), "I")
    s = Replace(s, ChrW(209), "N")
    s = Replace(s, ChrW(210), "O")
    s = Replace(s, ChrW(211), "O")
    s = Replace(s, ChrW(212), "O")
    s = Replace(s, ChrW(213), "O")
    s = Replace(s, ChrW(214), "O")
    s = Replace(s, ChrW(216), "O")
    s = Replace(s, ChrW(217), "U")
    s = Replace(s, ChrW(218), "U")
    s = Replace(s, ChrW(219), "U")
    s = Replace(s, ChrW(220), "U")
    s = Replace(s, ChrW(221), "Y")
    s = Replace(s, ChrW(173), vbNullString)
    s = Replace(s, ChrW(8211), "-")
    s = Replace(s, ChrW(8212), "-")
    s = Replace(s, vbTab, " ")

    NormalizarRotulo = LimparEspacos(s)
End Function

Private Function LimparSKU(ByVal valor As Variant) As String
    If IsError(valor) Or IsEmpty(valor) Or IsNull(valor) Then Exit Function

    LimparSKU = CStr(valor)
    LimparSKU = Replace(LimparSKU, ChrW(160), " ")
    LimparSKU = Replace(LimparSKU, ChrW(8203), vbNullString)
    LimparSKU = Replace(LimparSKU, ChrW(8204), vbNullString)
    LimparSKU = Replace(LimparSKU, ChrW(8205), vbNullString)
    LimparSKU = Replace(LimparSKU, ChrW(65279), vbNullString)
    LimparSKU = LimparEspacos(LimparSKU)
End Function

Private Function ChaveSKU(ByVal valor As Variant) As String
    Dim s As String

    s = UCase$(LimparSKU(valor))
    s = Replace(s, " ", vbNullString)
    s = Replace(s, vbTab, vbNullString)
    ChaveSKU = s
End Function

Private Function LimparEspacos(ByVal texto As String) As String
    Dim s As String

    s = Trim$(texto)
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    LimparEspacos = s
End Function

Private Function TextoSeguro(ByVal valor As Variant) As String
    If IsError(valor) Or IsEmpty(valor) Or IsNull(valor) Then
        TextoSeguro = vbNullString
    Else
        TextoSeguro = CStr(valor)
    End If
End Function

Private Function Contexto( _
    ByVal codigo As String, _
    ByVal campo As String) As String

    Contexto = "SKU " & codigo & " - " & campo
End Function

Private Sub AdicionarMotivo(ByRef lista As String, ByVal item As String)
    If Len(lista) = 0 Then
        lista = item
    Else
        lista = lista & "; " & item
    End If
End Sub

Private Function UltimaLinhaUsada(ByVal ws As Worksheet) As Long
    Dim ultimaCelula As Range

    Set ultimaCelula = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If ultimaCelula Is Nothing Then
        UltimaLinhaUsada = 0
    Else
        UltimaLinhaUsada = ultimaCelula.Row
    End If
End Function

Private Sub CapturarSnapshot(ByVal ws As Worksheet, ByRef snap As TSnapshot)
    Dim intervalo As Range

    Set intervalo = ws.UsedRange

    snap.PrimeiraLinha = intervalo.Row
    snap.PrimeiraColuna = intervalo.Column
    snap.NumeroLinhas = intervalo.Rows.Count
    snap.NumeroColunas = intervalo.Columns.Count
    snap.Conteudo = intervalo.Formula2
    snap.Valido = True
End Sub

Private Sub RestaurarSnapshot(ByVal ws As Worksheet, ByRef snap As TSnapshot)
    If ws Is Nothing Then Exit Sub
    If Not snap.Valido Then Exit Sub

    ws.UsedRange.ClearContents
    ws.Cells(snap.PrimeiraLinha, snap.PrimeiraColuna) _
      .Resize(snap.NumeroLinhas, snap.NumeroColunas).Formula2 = snap.Conteudo
End Sub





