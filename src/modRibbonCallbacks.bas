Attribute VB_Name = "modRibbonCallbacks"
'==============================================================================
' Módulo: modRibbonCallbacks
' Descrição: Callbacks para botões da Ribbon customizada
'==============================================================================
Option Explicit

Public myRibbon As IRibbonUI

' Callback para inicializar a ribbon
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
    modAPAFormatter.InicializarFormatter
End Sub

' Inserir Citação
Public Sub InserirCitacao(control As IRibbonControl)
    modAPAFormatter.InserirCitacao
End Sub

' Nova Referência
Public Sub NovaReferencia(control As IRibbonControl)
    Dim frm As frmNovaReferencia
    Set frm = New frmNovaReferencia
    frm.Show vbModal
End Sub

' Gerenciar Referências
Public Sub GerenciarReferencias(control As IRibbonControl)
    Dim frm As frmGerenciarReferencias
    Set frm = New frmGerenciarReferencias
    frm.Show vbModal
End Sub

' Gerar Bibliografia
Public Sub GerarListaReferencias(control As IRibbonControl)
    modAPAFormatter.GerarListaReferencias
End Sub

' Atualizar Citações
Public Sub AtualizarCitacoes(control As IRibbonControl)
    ' TODO: Implementar
    MsgBox "Funcionalidade em desenvolvimento", vbInformation
End Sub

' Validar Documento
Public Sub ValidarDocumento(control As IRibbonControl)
    modAPAFormatter.ValidarDocumento
End Sub

' Importar Referências
Public Sub ImportarReferencias(control As IRibbonControl)
    Dim arquivo As String
    
    arquivo = Application.FileDialog(msoFileDialogFilePicker).Show
    
    If Len(arquivo) > 0 Then
        ' TODO: Implementar importação
        MsgBox "Funcionalidade em desenvolvimento", vbInformation
    End If
End Sub

' Exportar Referências
Public Sub ExportarReferencias(control As IRibbonControl)
    ' TODO: Implementar exportação
    MsgBox "Funcionalidade em desenvolvimento", vbInformation
End Sub
