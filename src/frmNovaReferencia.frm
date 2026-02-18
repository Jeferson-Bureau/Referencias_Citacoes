VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNovaReferencia 
   Caption         =   "Nova Referência"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5000
   OleObjectBlob   =   "frmNovaReferencia.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNovaReferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' UserForm: frmNovaReferencia
' Descrição: Formulário para adicionar nova referência
'==============================================================================
Option Explicit

Private Sub UserForm_Initialize()
    ' Configurar tipo de referência
    cboTipoReferencia.AddItem "Livro"
    cboTipoReferencia.AddItem "Artigo de Periódico"
    cboTipoReferencia.AddItem "Site"
    cboTipoReferencia.AddItem "Tese/Dissertação"
    cboTipoReferencia.ListIndex = 0
End Sub

Private Sub cboTipoReferencia_Change()
    ' Mostrar/ocultar campos conforme tipo selecionado
    Select Case cboTipoReferencia.Text
        Case "Livro"
            lblEditora.Visible = True
            txtEditora.Visible = True
            lblEdicao.Visible = True
            txtEdicao.Visible = True
            
            lblPeriodico.Visible = False
            txtPeriodico.Visible = False
            lblVolume.Visible = False
            txtVolume.Visible = False
            
        Case "Artigo de Periódico"
            lblEditora.Visible = False
            txtEditora.Visible = False
            lblEdicao.Visible = False
            txtEdicao.Visible = False
            
            lblPeriodico.Visible = True
            txtPeriodico.Visible = True
            lblVolume.Visible = True
            txtVolume.Visible = True
    End Select
End Sub

Private Sub btnSalvar_Click()
    Dim ref As clsReferencia
    Dim autor As clsAutor
    
    ' Validar campos
    If Len(txtTitulo.Text) = 0 Then
        MsgBox "Título é obrigatório!", vbExclamation
        Exit Sub
    End If
    
    ' Criar referência
    Select Case cboTipoReferencia.Text
        Case "Livro"
            Dim livro As New clsReferenciaLivro
            livro.ID = "ref_" & Format(Now, "yyyymmddhhnnss")
            livro.Tipo = "livro"
            livro.Titulo = txtTitulo.Text
            livro.Ano = txtAno.Text
            livro.Editora = txtEditora.Text
            livro.Edicao = txtEdicao.Text
            livro.DOI = txtDOI.Text
            
            ' Adicionar autor
            Set autor = New clsAutor
            autor.Sobrenome = txtAutorSobrenome.Text
            autor.Iniciais = txtAutorIniciais.Text
            livro.AdicionarAutor autor
            
            ' Adicionar à coleção
            colReferencias.Add livro
            
            Set ref = livro
    End Select
    
    ' Validar
    Dim erros As String
    erros = ValidarReferencia(ref)
    
    If Len(erros) > 0 Then
        MsgBox "Erros na referência:" & vbCrLf & erros, vbExclamation
        Exit Sub
    End If
    
    ' Salvar no documento
    SalvarReferenciasNoDocumento
    
    MsgBox "Referência adicionada com sucesso!", vbInformation
    Me.Hide
End Sub
