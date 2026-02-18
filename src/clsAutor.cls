VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Módulo de Classe: clsAutor
' Descrição: Representa um autor de uma referência
'==============================================================================
Option Explicit

Private pSobrenome As String
Private pIniciais As String
Private pSufixo As String
Private pOrganizacao As Boolean

' Propriedades
Public Property Get Sobrenome() As String
    Sobrenome = pSobrenome
End Property

Public Property Let Sobrenome(ByVal valor As String)
    pSobrenome = Trim(valor)
End Property

Public Property Get Iniciais() As String
    Iniciais = pIniciais
End Property

Public Property Let Iniciais(ByVal valor As String)
    pIniciais = UCase(Trim(valor))
End Property

Public Property Get Sufixo() As String
    Sufixo = pSufixo
End Property

Public Property Let Sufixo(ByVal valor As String)
    pSufixo = Trim(valor)
End Property

Public Property Get Organizacao() As Boolean
    Organizacao = pOrganizacao
End Property

Public Property Let Organizacao(ByVal valor As Boolean)
    pOrganizacao = valor
End Property

' Método para formatar autor
Public Function FormatarAPA() As String
    If pOrganizacao Then
        FormatarAPA = pSobrenome
    Else
        Dim iniciaisFormatadas As String
        iniciaisFormatadas = FormatarIniciais(pIniciais)
        
        FormatarAPA = pSobrenome & ", " & iniciaisFormatadas
        
        If Len(pSufixo) > 0 Then
            FormatarAPA = FormatarAPA & ", " & pSufixo
        End If
    End If
End Function

Private Function FormatarIniciais(ByVal iniciais As String) As String
    Dim i As Integer
    Dim resultado As String
    
    ' Adicionar ponto e espaço após cada inicial
    For i = 1 To Len(iniciais)
        If Mid(iniciais, i, 1) <> " " And Mid(iniciais, i, 1) <> "." Then
            resultado = resultado & UCase(Mid(iniciais, i, 1)) & ". "
        End If
    Next i
    
    FormatarIniciais = Trim(resultado)
End Function
