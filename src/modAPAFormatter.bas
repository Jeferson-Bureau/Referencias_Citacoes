Attribute VB_Name = "modAPAFormatter"
'==============================================================================
' Módulo Padrão: modAPAFormatter
' Descrição: Funções principais de formatação APA 7
'==============================================================================
Option Explicit

' Variável global para armazenar referências
Public colReferencias As Collection

' Inicialização
Public Sub InicializarFormatter()
    Set colReferencias = New Collection
    CarregarReferenciasDoDocumento
End Sub

'==============================================================================
' INSERIR CITAÇÃO
'==============================================================================
Public Sub InserirCitacao()
    Dim frmCitacao As frmInserirCitacao
    Set frmCitacao = New frmInserirCitacao
    
    ' Mostrar formulário
    frmCitacao.Show vbModal
    
    ' Limpar
    Set frmCitacao = Nothing
End Sub

Public Sub InserirCitacaoNarrativa(ByVal referenciaID As String, _
                                   Optional ByVal pagina As String = "")
    Dim ref As clsReferencia
    Dim textoCitacao As String
    
    ' Buscar referência
    Set ref = BuscarReferencia(referenciaID)
    
    If ref Is Nothing Then
        MsgBox "Referência não encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Formatar citação narrativa
    textoCitacao = FormatarAutoresCitacao(ref.Autores) & " (" & ref.Ano
    
    If Len(pagina) > 0 Then
        textoCitacao = textoCitacao & ", p. " & pagina
    End If
    
    textoCitacao = textoCitacao & ")"
    
    ' Inserir no documento
    Selection.TypeText textoCitacao
End Sub

Public Sub InserirCitacaoParentetica(ByVal referenciaID As String, _
                                     Optional ByVal pagina As String = "")
    Dim ref As clsReferencia
    Dim textoCitacao As String
    
    ' Buscar referência
    Set ref = BuscarReferencia(referenciaID)
    
    If ref Is Nothing Then
        MsgBox "Referência não encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Formatar citação parentética
    textoCitacao = "(" & FormatarAutoresCitacao(ref.Autores) & ", " & ref.Ano
    
    If Len(pagina) > 0 Then
        textoCitacao = textoCitacao & ", p. " & pagina
    End If
    
    textoCitacao = textoCitacao & ")"
    
    ' Inserir no documento
    Selection.TypeText textoCitacao
End Sub

Private Function FormatarAutoresCitacao(ByRef autores As Collection) As String
    Dim autor As clsAutor
    Dim resultado As String
    
    If autores.Count = 0 Then
        FormatarAutoresCitacao = ""
        Exit Function
    End If
    
    ' 1 autor
    If autores.Count = 1 Then
        Set autor = autores(1)
        FormatarAutoresCitacao = autor.Sobrenome
        Exit Function
    End If
    
    ' 2 autores
    If autores.Count = 2 Then
        Set autor = autores(1)
        resultado = autor.Sobrenome
        
        Set autor = autores(2)
        resultado = resultado & " e " & autor.Sobrenome
        
        FormatarAutoresCitacao = resultado
        Exit Function
    End If
    
    ' 3+ autores
    Set autor = autores(1)
    FormatarAutoresCitacao = autor.Sobrenome & " et al."
End Function

'==============================================================================
' GERAR BIBLIOGRAFIA
'==============================================================================
Public Sub GerarListaReferencias()
    Dim doc As Document
    Dim rng As Range
    Dim ref As clsReferencia
    Dim i As Integer
    Dim refsOrdenadas As Collection
    
    Set doc = ActiveDocument
    
    ' Ordenar referências alfabeticamente
    Set refsOrdenadas = OrdenarReferenciasAlfabeticamente(colReferencias)
    
    ' Ir para o final do documento
    Set rng = doc.Content
    rng.Collapse Direction:=wdCollapseEnd
    
    ' Inserir título "Referências"
    rng.InsertAfter vbCrLf
    rng.Collapse Direction:=wdCollapseEnd
    rng.Style = wdStyleHeading1
    rng.Text = "Referências"
    rng.Collapse Direction:=wdCollapseEnd
    
    ' Inserir cada referência
    For Each ref In refsOrdenadas
        rng.InsertParagraphAfter
        Set rng = doc.Content
        rng.Collapse Direction:=wdCollapseEnd
        
        ' Inserir texto da referência
        rng.Text = ref.FormatarAPA
        
        ' Aplicar formatação
        With rng.ParagraphFormat
            .LeftIndent = CentimetersToPoints(1.27) ' 0.5 polegadas
            .FirstLineIndent = CentimetersToPoints(-1.27) ' Recuo pendente
            .LineSpacingRule = wdLineSpaceDouble ' Espaçamento duplo
            .SpaceAfter = 0
        End With
        
        rng.Collapse Direction:=wdCollapseEnd
    Next ref
    
    MsgBox "Bibliografia gerada com sucesso!", vbInformation
End Sub

'==============================================================================
' VALIDAÇÃO
'==============================================================================
Public Function ValidarReferencia(ByRef ref As clsReferencia) As String
    Dim erros As String
    
    ' Validar campos obrigatórios
    If Len(ref.ID) = 0 Then
        erros = erros & "- ID é obrigatório" & vbCrLf
    End If
    
    If Len(ref.Titulo) = 0 Then
        erros = erros & "- Título é obrigatório" & vbCrLf
    End If
    
    If ref.Autores.Count = 0 Then
        erros = erros & "- Ao menos um autor é necessário" & vbCrLf
    End If
    
    If Len(ref.Ano) = 0 Then
        erros = erros & "- Ano de publicação é obrigatório" & vbCrLf
    End If
    
    ' Validar DOI se presente
    If Len(ref.DOI) > 0 Then
        If Not ValidarFormatoDOI(ref.DOI) Then
            erros = erros & "- Formato de DOI inválido" & vbCrLf
        End If
    End If
    
    ValidarReferencia = erros
End Function

Private Function ValidarFormatoDOI(ByVal doi As String) As Boolean
    ' Verificar se começa com 10.
    If Left(doi, 3) = "10." Then
        ValidarFormatoDOI = True
    Else
        ValidarFormatoDOI = False
    End If
End Function

Public Sub ValidarDocumento()
    Dim ref As clsReferencia
    Dim erros As String
    Dim totalErros As Integer
    
    totalErros = 0
    
    For Each ref In colReferencias
        Dim validacao As String
        validacao = ValidarReferencia(ref)
        
        If Len(validacao) > 0 Then
            totalErros = totalErros + 1
            erros = erros & "Referência: " & ref.Titulo & vbCrLf
            erros = erros & validacao & vbCrLf
        End If
    Next ref
    
    If totalErros = 0 Then
        MsgBox "Todas as referências estão válidas!", vbInformation
    Else
        MsgBox "Erros encontrados:" & vbCrLf & vbCrLf & erros, vbExclamation
    End If
End Sub

'==============================================================================
' UTILITÁRIOS
'==============================================================================
Private Function BuscarReferencia(ByVal ID As String) As clsReferencia
    Dim ref As clsReferencia
    
    For Each ref In colReferencias
        If ref.ID = ID Then
            Set BuscarReferencia = ref
            Exit Function
        End If
    Next ref
    
    Set BuscarReferencia = Nothing
End Function

Private Function OrdenarReferenciasAlfabeticamente(ByRef refs As Collection) As Collection
    Dim resultado As New Collection
    Dim ref As clsReferencia
    Dim i As Integer
    Dim j As Integer
    Dim temp As clsReferencia
    Dim arrRefs() As clsReferencia
    
    ' Converter collection para array
    ReDim arrRefs(1 To refs.Count)
    
    i = 1
    For Each ref In refs
        Set arrRefs(i) = ref
        i = i + 1
    Next ref
    
    ' Bubble sort
    For i = 1 To UBound(arrRefs) - 1
        For j = i + 1 To UBound(arrRefs)
            Dim sobrenome1 As String
            Dim sobrenome2 As String
            
            If arrRefs(i).Autores.Count > 0 Then
                sobrenome1 = arrRefs(i).Autores(1).Sobrenome
            Else
                sobrenome1 = arrRefs(i).Titulo
            End If
            
            If arrRefs(j).Autores.Count > 0 Then
                sobrenome2 = arrRefs(j).Autores(1).Sobrenome
            Else
                sobrenome2 = arrRefs(j).Titulo
            End If
            
            If StrComp(sobrenome1, sobrenome2, vbTextCompare) > 0 Then
                Set temp = arrRefs(i)
                Set arrRefs(i) = arrRefs(j)
                Set arrRefs(j) = temp
            End If
        Next j
    Next i
    
    ' Converter array de volta para collection
    For i = 1 To UBound(arrRefs)
        resultado.Add arrRefs(i)
    Next i
    
    Set OrdenarReferenciasAlfabeticamente = resultado
End Function

'==============================================================================
' PERSISTÊNCIA DE DADOS
'==============================================================================
Public Sub SalvarReferenciasNoDocumento()
    Dim doc As Document
    Dim xmlPart As CustomXMLPart
    Dim xmlData As String
    Dim ref As clsReferencia
    
    Set doc = ActiveDocument
    
    ' Construir XML
    xmlData = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    xmlData = xmlData & "<referencias>" & vbCrLf
    
    For Each ref In colReferencias
        xmlData = xmlData & "  <referencia>" & vbCrLf
        xmlData = xmlData & "    <id>" & EscapeXML(ref.ID) & "</id>" & vbCrLf
        xmlData = xmlData & "    <tipo>" & EscapeXML(ref.Tipo) & "</tipo>" & vbCrLf
        xmlData = xmlData & "    <titulo>" & EscapeXML(ref.Titulo) & "</titulo>" & vbCrLf
        xmlData = xmlData & "    <ano>" & EscapeXML(ref.Ano) & "</ano>" & vbCrLf
        ' ... outros campos
        xmlData = xmlData & "  </referencia>" & vbCrLf
    Next ref
    
    xmlData = xmlData & "</referencias>"
    
    ' Salvar no documento
    On Error Resume Next
    doc.CustomXMLParts.Add xmlData
    On Error GoTo 0
End Sub

Public Sub CarregarReferenciasDoDocumento()
    Dim doc As Document
    Dim xmlPart As CustomXMLPart
    Dim xmlDoc As Object
    Dim nodeList As Object
    Dim node As Object
    Dim ref As clsReferencia
    
    Set doc = ActiveDocument
    Set colReferencias = New Collection
    
    ' TODO: Implementar parser XML
    ' Por simplicidade, pode-se usar Custom Document Properties
    ' ou um arquivo externo
End Sub

Private Function EscapeXML(ByVal texto As String) As String
    texto = Replace(texto, "&", "&amp;")
    texto = Replace(texto, "<", "&lt;")
    texto = Replace(texto, ">", "&gt;")
    texto = Replace(texto, """", "&quot;")
    texto = Replace(texto, "'", "&apos;")
    EscapeXML = texto
End Function

'==============================================================================
' FUNÇÕES DE FORMATAÇÃO DE TEXTO
'==============================================================================
Public Function CapitalizarTituloAPA(ByVal titulo As String) As String
    Dim palavras() As String
    Dim i As Integer
    Dim resultado As String
    
    palavras = Split(titulo, " ")
    
    For i = LBound(palavras) To UBound(palavras)
        If i = LBound(palavras) Then
            ' Primeira palavra sempre maiúscula
            resultado = UCase(Left(palavras(i), 1)) & LCase(Mid(palavras(i), 2))
        Else
            ' Demais palavras em minúscula
            resultado = resultado & " " & LCase(palavras(i))
        End If
    Next i
    
    CapitalizarTituloAPA = resultado
End Function
