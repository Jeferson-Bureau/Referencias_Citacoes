VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInserirCitacao 
   Caption         =   "Inserir Citação APA"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "frmInserirCitacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInserirCitacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' UserForm: frmInserirCitacao
' Descrição: Formulário para inserir citação no documento
'==============================================================================
Option Explicit

Private Sub UserForm_Initialize()
    ' Carregar lista de referências
    Dim ref As clsReferencia
    
    lstReferencias.Clear
    
    For Each ref In colReferencias
        Dim item As String
        item = ref.Autores(1).Sobrenome & " (" & ref.Ano & ") - " & ref.Titulo
        lstReferencias.AddItem item
        ' Armazenar ID na segunda coluna (invisível ou auxiliar)
        ' Nota: ListBox multi-coluna requer configuração de ColumnCount
        lstReferencias.List(lstReferencias.ListCount - 1, 1) = ref.ID
    Next ref
    
    ' Configurar tipo de citação
    cboTipoCitacao.AddItem "Narrativa"
    cboTipoCitacao.AddItem "Parentética"
    cboTipoCitacao.ListIndex = 1
End Sub

Private Sub btnInserir_Click()
    If lstReferencias.ListIndex = -1 Then
        MsgBox "Selecione uma referência!", vbExclamation
        Exit Sub
    End If
    
    Dim referenciaID As String
    Dim tipoCitacao As String
    Dim pagina As String
    
    referenciaID = lstReferencias.List(lstReferencias.ListIndex, 1)
    tipoCitacao = cboTipoCitacao.Text
    pagina = txtPagina.Text
    
    If tipoCitacao = "Narrativa" Then
        InserirCitacaoNarrativa referenciaID, pagina
    Else
        InserirCitacaoParentetica referenciaID, pagina
    End If
    
    Me.Hide
End Sub

Private Sub btnCancelar_Click()
    Me.Hide
End Sub
