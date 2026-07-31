Attribute VB_Name = "InstalarAutomacao_ORIGINAL"
' =============================================================================
'  CODIGO ORIGINAL ENVIADO PELO USUARIO - mantido apenas para referencia.
'  A versao corrigida esta em InstalarAutomacao.bas
' =============================================================================
Option Explicit

Private Const PASTA_ABERTAS As String = "C:\Users\CEmerson\Ultradent Products Inc\PCP e Logística - Hard work - Documentos\CONTROL BOARD\EXTRACOES\OP_ABERTAS"
Private Const PASTA_FECHADAS As String = "C:\Users\CEmerson\Ultradent Products Inc\PCP e Logística - Hard work - Documentos\CONTROL BOARD\EXTRACOES\OP_FECHADAS"
Private Const NOME_BASE As String = "_BASE_OP"

Public Sub InstalarAutomacao()
    On Error GoTo TrataErro
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Instalando automação da Programação..."

    RemoverConsultasIncompletas
    ConfigurarListaStatus
    AtualizarProgramacao

    ThisWorkbook.Worksheets("INSTALAR_AUTOMACAO").Visible = xlSheetVeryHidden
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Automação instalada e base atualizada com sucesso.", vbInformation
    Exit Sub

TrataErro:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Falha na etapa de instalação: " & Err.Description, vbExclamation
End Sub

' (demais rotinas identicas ao anexo enviado - omitidas aqui de proposito;
'  consulte a versao corrigida InstalarAutomacao.bas)
