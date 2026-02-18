# Prompt para Desenvolvimento de Macro Word - Formatação APA 7 (VBA)

## Contexto do Projeto

Você é um desenvolvedor VBA (Visual Basic for Applications) sênior especializado em automação de documentos Microsoft Word. Sua missão é criar um conjunto de macros e módulos VBA que automatizem a formatação de referências bibliográficas e citações no padrão APA 7ª edição, integrado diretamente ao Microsoft Word.

## Requisitos Técnicos

### Stack Tecnológica

```vba
' Tecnologias principais:
- VBA (Visual Basic for Applications) 7.1+
- Microsoft Word Object Model
- Word UserForms (interface gráfica)
- Collections e Dictionaries (estruturas de dados)
- XML para persistência de dados
- Ribbon XML (customização da interface)
```

### Ambiente de Desenvolvimento

- **Microsoft Word** 2016 ou superior
- **Editor VBA** (Alt + F11)
- **Referências necessárias**:
  - Microsoft Word Object Library
  - Microsoft Scripting Runtime (Dictionary)
  - Microsoft XML v6.0

## Especificações Funcionais

### 1. Estrutura de Classes VBA

```vba
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

' Método abstrato - deve ser sobrescrito nas classes filhas
Public Function FormatarAPA() As String
    FormatarAPA = ""
End Function

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
```

### 2. Módulo Principal de Formatação

```vba
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
```

### 3. UserForms (Interface Gráfica)

```vba
'==============================================================================
' UserForm: frmInserirCitacao
' Descrição: Formulário para inserir citação no documento
'==============================================================================

' Código do formulário
Private Sub UserForm_Initialize()
    ' Carregar lista de referências
    Dim ref As clsReferencia
    
    lstReferencias.Clear
    
    For Each ref In colReferencias
        Dim item As String
        item = ref.Autores(1).Sobrenome & " (" & ref.Ano & ") - " & ref.Titulo
        lstReferencias.AddItem item
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

'==============================================================================
' UserForm: frmNovaReferencia
' Descrição: Formulário para adicionar nova referência
'==============================================================================

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

'==============================================================================
' UserForm: frmGerenciarReferencias
' Descrição: Formulário para gerenciar referências existentes
'==============================================================================

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
```

### 4. Customização da Ribbon

```xml
<!-- Arquivo CustomUI.xml -->
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabAPAFormatter" label="APA 7">
        
        <!-- Grupo: Citações -->
        <group id="grpCitacoes" label="Citações">
          <button id="btnInserirCitacao"
                  label="Inserir Citação"
                  size="large"
                  imageMso="QuoteMarksInsert"
                  onAction="InserirCitacao"
                  screentip="Inserir Citação"
                  supertip="Insere uma citação no texto na posição do cursor"/>
        </group>
        
        <!-- Grupo: Referências -->
        <group id="grpReferencias" label="Referências">
          <button id="btnNovaReferencia"
                  label="Nova Referência"
                  size="large"
                  imageMso="RecordsAddFromOutlook"
                  onAction="NovaReferencia"
                  screentip="Nova Referência"
                  supertip="Adiciona uma nova referência bibliográfica"/>
                  
          <button id="btnGerenciarReferencias"
                  label="Gerenciar"
                  size="large"
                  imageMso="DatabaseProperties"
                  onAction="GerenciarReferencias"
                  screentip="Gerenciar Referências"
                  supertip="Visualiza e edita referências existentes"/>
        </group>
        
        <!-- Grupo: Bibliografia -->
        <group id="grpBibliografia" label="Bibliografia">
          <button id="btnGerarBibliografia"
                  label="Gerar Bibliografia"
                  size="large"
                  imageMso="TableOfContentsInsertGallery"
                  onAction="GerarListaReferencias"
                  screentip="Gerar Bibliografia"
                  supertip="Gera a lista de referências formatada em APA 7"/>
                  
          <button id="btnAtualizarCitacoes"
                  label="Atualizar Citações"
                  size="normal"
                  imageMso="RefreshArrows"
                  onAction="AtualizarCitacoes"
                  screentip="Atualizar Citações"
                  supertip="Atualiza todas as citações no documento"/>
        </group>
        
        <!-- Grupo: Validação -->
        <group id="grpValidacao" label="Validação">
          <button id="btnValidar"
                  label="Validar"
                  size="large"
                  imageMso="AcceptInvitation"
                  onAction="ValidarDocumento"
                  screentip="Validar Documento"
                  supertip="Valida todas as referências e citações"/>
        </group>
        
        <!-- Grupo: Importar/Exportar -->
        <group id="grpImportExport" label="Importar/Exportar">
          <button id="btnImportar"
                  label="Importar"
                  size="normal"
                  imageMso="FileOpen"
                  onAction="ImportarReferencias"
                  screentip="Importar"
                  supertip="Importa referências de arquivo BibTeX ou RIS"/>
                  
          <button id="btnExportar"
                  label="Exportar"
                  size="normal"
                  imageMso="FileSaveAs"
                  onAction="ExportarReferencias"
                  screentip="Exportar"
                  supertip="Exporta referências para arquivo"/>
        </group>
        
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

### 5. Callbacks da Ribbon

```vba
'==============================================================================
' Módulo: modRibbonCallbacks
' Descrição: Callbacks para botões da Ribbon customizada
'==============================================================================
Option Explicit

' Callback para inicializar a ribbon
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
    InicializarFormatter
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
```

## Estrutura do Projeto VBA

```
FormatadorAPA7.docm (Documento Word com Macros)
│
├── ThisDocument (Módulo do Documento)
│   └── Eventos do documento
│
├── Módulos de Classe
│   ├── clsAutor
│   ├── clsReferencia
│   ├── clsReferenciaLivro
│   ├── clsReferenciaArtigo
│   ├── clsReferenciaSite
│   └── ... (outras 8 classes de referência)
│
├── Módulos Padrão
│   ├── modAPAFormatter (Principal)
│   ├── modUtilitarios (Funções auxiliares)
│   ├── modValidacao (Validações)
│   ├── modPersistencia (Salvar/carregar)
│   └── modRibbonCallbacks (Callbacks da Ribbon)
│
├── UserForms
│   ├── frmInserirCitacao
│   ├── frmNovaReferencia
│   ├── frmGerenciarReferencias
│   ├── frmValidacao
│   └── frmConfiguracao
│
└── CustomUI (XML da Ribbon)
    └── CustomUI.xml
```

## Instalação e Configuração

### Passo 1: Criar Documento Base

1. Abra o Microsoft Word
2. Salve como: `FormatadorAPA7.docm` (Documento Word Habilitado para Macro)
3. Pressione `Alt + F11` para abrir o Editor VBA

### Passo 2: Adicionar Referências

No Editor VBA:

1. Menu: **Ferramentas** > **Referências**
2. Marcar:
   - Microsoft Word Object Library
   - Microsoft Scripting Runtime
   - Microsoft XML, v6.0

### Passo 3: Criar Módulos e Classes

1. **Inserir** > **Módulo de Classe**
   - Criar: clsAutor, clsReferencia, clsReferenciaLivro, etc.

2. **Inserir** > **Módulo**
   - Criar: modAPAFormatter, modUtilitarios, etc.

3. **Inserir** > **UserForm**
   - Criar: frmInserirCitacao, frmNovaReferencia, etc.

### Passo 4: Customizar Ribbon (Opcional)

1. Instalar **Office RibbonX Editor**
2. Abrir `FormatadorAPA7.docm` no editor
3. Adicionar XML customizado
4. Salvar

## Exemplos de Uso

### Exemplo 1: Criar e Adicionar Referência de Livro

```vba
Sub ExemploAdicionarLivro()
    Dim livro As New clsReferenciaLivro
    Dim autor As New clsAutor
    
    ' Configurar autor
    autor.Sobrenome = "Silva"
    autor.Iniciais = "J M"
    
    ' Configurar livro
    livro.ID = "silva2020intro"
    livro.Titulo = "Introdução à programação"
    livro.Ano = "2020"
    livro.Editora = "Tech Books"
    livro.Edicao = "3ª ed."
    livro.DOI = "10.1234/example"
    
    ' Adicionar autor ao livro
    livro.AdicionarAutor autor
    
    ' Adicionar à coleção global
    colReferencias.Add livro
    
    ' Salvar no documento
    SalvarReferenciasNoDocumento
    
    MsgBox "Livro adicionado: " & livro.FormatarAPA
End Sub
```

### Exemplo 2: Inserir Citação Programaticamente

```vba
Sub ExemploInserirCitacao()
    ' Garantir que o formatador está inicializado
    InicializarFormatter
    
    ' Inserir citação narrativa
    InserirCitacaoNarrativa "silva2020intro"
    ' Resultado no documento: Silva (2020)
    
    ' Inserir citação parentética com página
    Selection.TypeText " afirma que VBA é poderoso "
    InserirCitacaoParentetica "silva2020intro", "45"
    ' Resultado no documento: (Silva, 2020, p. 45)
End Sub
```

### Exemplo 3: Gerar Bibliografia Completa

```vba
Sub ExemploGerarBibliografia()
    ' Garantir que há referências
    If colReferencias.Count = 0 Then
        MsgBox "Nenhuma referência cadastrada!", vbExclamation
        Exit Sub
    End If
    
    ' Gerar bibliografia
    GerarListaReferencias
    
    ' Resultado: Seção "Referências" no final do documento
    ' com todas as referências formatadas em APA 7
End Sub
```

### Exemplo 4: Validar Todas as Referências

```vba
Sub ExemploValidar()
    Dim ref As clsReferencia
    Dim totalErros As Integer
    Dim relatorio As String
    
    totalErros = 0
    relatorio = "RELATÓRIO DE VALIDAÇÃO" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    
    For Each ref In colReferencias
        Dim erros As String
        erros = ValidarReferencia(ref)
        
        If Len(erros) > 0 Then
            totalErros = totalErros + 1
            relatorio = relatorio & "Referência: " & ref.Titulo & vbCrLf
            relatorio = relatorio & erros & vbCrLf
        End If
    Next ref
    
    If totalErros = 0 Then
        MsgBox "✓ Todas as " & colReferencias.Count & " referências estão válidas!", vbInformation
    Else
        MsgBox relatorio, vbExclamation, "Erros de Validação"
    End If
End Sub
```

## Funcionalidades Avançadas

### Auto_Open e Auto_Close

```vba
'==============================================================================
' ThisDocument (Módulo do Documento)
'==============================================================================

Private Sub Document_Open()
    ' Inicializar ao abrir documento
    InicializarFormatter
End Sub

Private Sub Document_Close()
    ' Salvar ao fechar documento
    SalvarReferenciasNoDocumento
End Sub
```

### Atalhos de Teclado

```vba
Sub ConfigurarAtalhos()
    ' Ctrl+Shift+C = Inserir Citação
    Application.CustomizationContext = ActiveDocument
    Application.KeyBindings.Add _
        KeyCategory:=wdKeyCategoryCommand, _
        Command:="InserirCitacao", _
        KeyCode:=BuildKeyCode(wdKeyC, wdKeyControl, wdKeyShift)
    
    ' Ctrl+Shift+B = Gerar Bibliografia
    Application.KeyBindings.Add _
        KeyCategory:=wdKeyCategoryCommand, _
        Command:="GerarListaReferencias", _
        KeyCode:=BuildKeyCode(wdKeyB, wdKeyControl, wdKeyShift)
End Sub
```

### Importar de BibTeX (Básico)

```vba
Sub ImportarBibTeX()
    Dim arquivo As String
    Dim textoCompleto As String
    Dim linhas() As String
    Dim i As Integer
    
    ' Selecionar arquivo
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Selecione arquivo BibTeX"
        .Filters.Add "BibTeX", "*.bib"
        
        If .Show = -1 Then
            arquivo = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Ler arquivo
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(arquivo, 1) ' 1 = ForReading
    
    textoCompleto = ts.ReadAll
    ts.Close
    
    ' Parsear (implementação simplificada)
    linhas = Split(textoCompleto, vbCrLf)
    
    ' TODO: Implementar parser completo
    MsgBox "Arquivo lido. " & UBound(linhas) & " linhas encontradas.", vbInformation
End Sub
```

### Exportar para JSON

```vba
Sub ExportarParaJSON()
    Dim json As String
    Dim ref As clsReferencia
    Dim arquivo As String
    
    ' Construir JSON
    json = "{" & vbCrLf
    json = json & "  ""referencias"": [" & vbCrLf
    
    For Each ref In colReferencias
        json = json & "    {" & vbCrLf
        json = json & "      ""id"": """ & ref.ID & """," & vbCrLf
        json = json & "      ""tipo"": """ & ref.Tipo & """," & vbCrLf
        json = json & "      ""titulo"": """ & EscapeJSON(ref.Titulo) & """," & vbCrLf
        json = json & "      ""ano"": """ & ref.Ano & """" & vbCrLf
        json = json & "    }," & vbCrLf
    Next ref
    
    ' Remover última vírgula
    json = Left(json, Len(json) - 3) & vbCrLf
    
    json = json & "  ]" & vbCrLf
    json = json & "}"
    
    ' Salvar arquivo
    arquivo = Application.ActiveDocument.Path & "\referencias.json"
    
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(arquivo, True)
    ts.Write json
    ts.Close
    
    MsgBox "Referências exportadas para: " & arquivo, vbInformation
End Sub

Private Function EscapeJSON(ByVal texto As String) As String
    texto = Replace(texto, "\", "\\")
    texto = Replace(texto, """", "\""")
    texto = Replace(texto, vbCrLf, "\n")
    texto = Replace(texto, vbCr, "\r")
    texto = Replace(texto, vbTab, "\t")
    EscapeJSON = texto
End Function
```

## Checklist de Implementação

### Fase 1: Estrutura Base ✓

- [ ] Criar documento .docm
- [ ] Configurar referências VBA
- [ ] Criar estrutura de módulos e classes
- [ ] Implementar clsAutor
- [ ] Implementar clsReferencia (base)

### Fase 2: Classes de Referência ✓

- [ ] clsReferenciaLivro
- [ ] clsReferenciaArtigo
- [ ] clsReferenciaSite
- [ ] clsReferenciaTese
- [ ] ... (outros 7 tipos)

### Fase 3: Formatadores ✓

- [ ] Função FormatarAutor
- [ ] Função FormatarListaAutores
- [ ] Função FormatarAutoresCitacao
- [ ] Função CapitalizarTituloAPA
- [ ] Método FormatarAPA para cada classe

### Fase 4: Citações ✓

- [ ] InserirCitacaoNarrativa
- [ ] InserirCitacaoParentetica
- [ ] InserirCitacaoDireta
- [ ] Suporte para múltiplas citações

### Fase 5: Bibliografia ✓

- [ ] GerarListaReferencias
- [ ] OrdenarReferenciasAlfabeticamente
- [ ] Aplicar formatação (recuo, espaçamento)
- [ ] Inserir em seção específica

### Fase 6: Validação ✓

- [ ] ValidarReferencia (individual)
- [ ] ValidarDocumento (todas)
- [ ] Detectar duplicatas
- [ ] Verificar citações órfãs

### Fase 7: Persistência ✓

- [ ] SalvarReferenciasNoDocumento
- [ ] CarregarReferenciasDoDocumento
- [ ] Usar Custom XML Parts ou Document Properties

### Fase 8: Interface ✓

- [ ] frmInserirCitacao
- [ ] frmNovaReferencia (com campos dinâmicos)
- [ ] frmGerenciarReferencias (lista, editar, excluir)
- [ ] frmValidacao (relatório)
- [ ] frmConfiguracao

### Fase 9: Ribbon Customizada ✓

- [ ] Criar CustomUI.xml
- [ ] Adicionar aba "APA 7"
- [ ] Botões principais
- [ ] Callbacks

### Fase 10: Extras ✓

- [ ] Importar BibTeX
- [ ] Importar RIS
- [ ] Exportar JSON
- [ ] Atalhos de teclado
- [ ] Auto_Open / Auto_Close

## Dicas e Melhores Práticas

### Performance

```vba
' Desabilitar atualização de tela durante operações longas
Sub ExemploPerformance()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ... código aqui ...
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
```

### Tratamento de Erros

```vba
Sub ExemploErros()
    On Error GoTo TratarErro
    
    ' ... código aqui ...
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro: " & Err.Description, vbCritical
    Err.Clear
End Sub
```

### Debug

```vba
' Use Debug.Print para log no Immediate Window
Debug.Print "Valor da variável: " & minhaVariavel

' Use breakpoints (F9) para pausar execução
' Use Step Into (F8) para executar linha por linha
```

## Recursos de Referência

- **APA Style Manual (7th edition):** <https://apastyle.apa.org/>
- **Word VBA Reference:** <https://docs.microsoft.com/office/vba/api/overview/word>
- **Ribbon XML Reference:** <https://docs.microsoft.com/openspecs/office_standards/>
- **VBA Best Practices:** <https://rubberduckvba.com/>

## Notas Importantes

1. **Sempre teste** em documento de teste antes de usar em trabalho real
2. **Faça backup** do documento antes de executar macros
3. **Valide dados** do usuário antes de processar
4. **Use Option Explicit** em todos os módulos
5. **Comente** o código adequadamente
6. **Trate erros** com On Error GoTo
7. **Desabilite ScreenUpdating** em operações longas
8. **Teste compatibilidade** com diferentes versões do Word

---

**Implemente este sistema VBA seguindo as melhores práticas de desenvolvimento e garantindo conformidade rigorosa com o Manual APA 7ª edição.**
