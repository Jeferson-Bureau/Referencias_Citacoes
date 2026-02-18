VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Módulo de Classe: clsReferencia
' Descrição: Classe base para todas as referências
'==============================================================================
Option Explicit

Private pID As String
Private pTipo As String
Private pAutores As Collection
Private pAno As String
Private pMes As String
Private pDia As String
Private pSemData As Boolean
Private pTitulo As String
Private pDOI As String
Private pURL As String
Private pIdioma As String

' Inicialização
Private Sub Class_Initialize()
    Set pAutores = New Collection
    pIdioma = "pt-BR"
End Sub

' Propriedades
Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal valor As String)
    pID = valor
End Property

Public Property Get Tipo() As String
    Tipo = pTipo
End Property

Public Property Let Tipo(ByVal valor As String)
    pTipo = valor
End Property

Public Property Get Autores() As Collection
    Set Autores = pAutores
End Property

Public Sub AdicionarAutor(ByRef autor As clsAutor)
    pAutores.Add autor
End Sub

Public Property Get Ano() As String
    Ano = pAno
End Property

Public Property Let Ano(ByVal valor As String)
    pAno = valor
End Property

Public Property Get Titulo() As String
    Titulo = pTitulo
End Property

Public Property Let Titulo(ByVal valor As String)
    pTitulo = valor
End Property

Public Property Get DOI() As String
    DOI = pDOI
End Property

Public Property Let DOI(ByVal valor As String)
    pDOI = valor
End Property

Public Property Get URL() As String
    URL = pURL
End Property

Public Property Let URL(ByVal valor As String)
    pURL = valor
End Property


' Método abstrato - deve ser sobrescrito nas classes filhas
Public Function FormatarAPA() As String
    FormatarAPA = ""
End Function
