VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReferenciaLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Módulo de Classe: clsReferenciaLivro
' Descrição: Referência específica para livros
'==============================================================================
Option Explicit

Private pEditora As String
Private pEdicao As String
Private pLocalPublicacao As String

Public Property Get Editora() As String
    Editora = pEditora
End Property

Public Property Let Editora(ByVal valor As String)
    pEditora = valor
End Property

Public Property Get Edicao() As String
    Edicao = pEdicao
End Property

Public Property Let Edicao(ByVal valor As String)
    pEdicao = valor
End Property

Public Function FormatarAPA() As String
    Dim resultado As String
    Dim i As Integer
    
    ' Formatar autores
    resultado = FormatarListaAutores(Me.Autores)
    
    ' Data
    resultado = resultado & " (" & Me.Ano & "). "
    
    ' Título em itálico
    resultado = resultado & Me.Titulo
    
    ' Edição
    If Len(pEdicao) > 0 Then
        resultado = resultado & " (" & pEdicao & ")"
    End If
    
    resultado = resultado & ". "
    
    ' Editora
    resultado = resultado & pEditora & "."
    
    ' DOI
    If Len(Me.DOI) > 0 Then
        resultado = resultado & " https://doi.org/" & Me.DOI
    ElseIf Len(Me.URL) > 0 Then
        resultado = resultado & " " & Me.URL
    End If
    
    FormatarAPA = resultado
End Function

Private Function FormatarListaAutores(ByRef autores As Collection) As String
    Dim resultado As String
    Dim i As Integer
    Dim autor As clsAutor
    Dim total As Integer
    
    total = autores.Count
    
    If total = 0 Then
        FormatarListaAutores = ""
        Exit Function
    End If
    
    ' 1 autor
    If total = 1 Then
        Set autor = autores(1)
        FormatarListaAutores = autor.FormatarAPA
        Exit Function
    End If
    
    ' 2 autores
    If total = 2 Then
        Set autor = autores(1)
        resultado = autor.FormatarAPA
        
        Set autor = autores(2)
        resultado = resultado & ", & " & autor.FormatarAPA
        
        FormatarListaAutores = resultado
        Exit Function
    End If
    
    ' 3-20 autores
    If total >= 3 And total <= 20 Then
        For i = 1 To total - 1
            Set autor = autores(i)
            resultado = resultado & autor.FormatarAPA & ", "
        Next i
        
        Set autor = autores(total)
        resultado = resultado & "& " & autor.FormatarAPA
        
        FormatarListaAutores = resultado
        Exit Function
    End If
    
    ' 21+ autores
    If total >= 21 Then
        For i = 1 To 19
            Set autor = autores(i)
            resultado = resultado & autor.FormatarAPA & ", "
        Next i
        
        resultado = resultado & "... "
        
        Set autor = autores(total)
        resultado = resultado & autor.FormatarAPA
        
        FormatarListaAutores = resultado
        Exit Function
    End If
End Function
