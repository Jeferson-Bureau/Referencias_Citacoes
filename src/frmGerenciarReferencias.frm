VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGerenciarReferencias 
   Caption         =   "Gerenciar Referências"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "frmGerenciarReferencias.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGerenciarReferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' UserForm: frmGerenciarReferencias
' Descrição: Formulário para gerenciar referências existentes
'==============================================================================
Option Explicit

Private Sub UserForm_Initialize()
    AtualizarLista
End Sub

Private Sub AtualizarLista()
    Dim ref As clsReferencia
    
    lstReferencias.Clear
    
    For Each ref In colReferencias
        Dim item As String
        item = ref.Autores(1).Sobrenome & " (" & ref.Ano & ")"
        lstReferencias.AddItem item
        lstReferencias.List(lstReferencias.ListCount - 1, 1) = ref.Titulo
        lstReferencias.List(lstReferencias.ListCount - 1, 2) = ref.ID
    Next ref
End Sub

Private Sub btnEditar_Click()
    If lstReferencias.ListIndex = -1 Then
        MsgBox "Selecione uma referência!", vbExclamation
        Exit Sub
    End If
    
    ' Abrir formulário de edição
    ' TODO: Implementar
End Sub

Private Sub btnExcluir_Click()
    If lstReferencias.ListIndex = -1 Then
        MsgBox "Selecione uma referência!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir esta referência?", vbYesNo + vbQuestion) = vbYes Then
        Dim referenciaID As String
        referenciaID = lstReferencias.List(lstReferencias.ListIndex, 2)
        
        ' Remover da coleção
        Dim i As Integer
        For i = 1 To colReferencias.Count
            If colReferencias(i).ID = referenciaID Then
                colReferencias.Remove i
                Exit For
            End If
        Next i
        
        AtualizarLista
        SalvarReferenciasNoDocumento
    End If
End Sub

Private Sub txtBuscar_Change()
    Dim ref As clsReferencia
    Dim termoBusca As String
    
    termoBusca = LCase(txtBuscar.Text)
    lstReferencias.Clear
    
    If Len(termoBusca) = 0 Then
        AtualizarLista
        Exit Sub
    End If
    
    For Each ref In colReferencias
        Dim textoCompleto As String
        textoCompleto = LCase(ref.Autores(1).Sobrenome & " " & ref.Titulo & " " & ref.Ano)
        
        If InStr(textoCompleto, termoBusca) > 0 Then
            Dim item As String
            item = ref.Autores(1).Sobrenome & " (" & ref.Ano & ")"
            lstReferencias.AddItem item
            lstReferencias.List(lstReferencias.ListCount - 1, 1) = ref.Titulo
            lstReferencias.List(lstReferencias.ListCount - 1, 2) = ref.ID
        End If
    Next ref
End Sub
