# Instruções de Instalação do Projeto VBA APA 7

Os arquivos fonte do projeto VBA foram gerados na pasta `src/`. Siga os passos abaixo para importar o projeto para o Microsoft Word.

## 1. Preparar o Documento Word

1. Abra o Microsoft Word.
2. Crie um novo documento ou abra o documento onde deseja instalar a macro.
3. Salve o arquivo como **Documento Habilitado para Macro do Word (.docm)** via `Arquivo > Salvar Como`.

## 2. Acessar o Editor VBA

1. Pressione `Alt + F11` para abrir o Editor do Visual Basic for Applications (VBA).

## 3. Adicionar Referências Necessárias

O projeto depende de bibliotecas externas. No Editor VBA:

1. Vá em **Ferramentas (Tools) > Referências (References)**.
2. Localize e marque as seguintes caixas:
    * [x] **Microsoft Word Object Library** (Geralmente já vem marcado)
    * [ ] **Microsoft Scripting Runtime**
    * [ ] **Microsoft XML, v6.0**
3. Clique em **OK**.

## 4. Importar os Arquivos

1. No Editor VBA, clique com o botão direito no projeto (Ex: `Project (FormatadorAPA7)`).
2. Selecione **Importar Arquivo (Import File...)**.
3. Navegue até a pasta `src/` deste projeto.
4. Selecione e importe todos os arquivos `.cls`, `.bas` e `.frm` um por um:
    * `clsAutor.cls`
    * `clsReferencia.cls`
    * `clsReferenciaLivro.cls`
    * `modAPAFormatter.bas`
    * `modRibbonCallbacks.bas`
    * `frmInserirCitacao.frm`
    * `frmNovaReferencia.frm`
    * `frmGerenciarReferencias.frm`

## 5. Configurar a Ribbon (Opcional, Avançado)

Para adicionar os botões na barra de ferramentas do Word:

1. Você precisará de uma ferramenta externa como o **Office RibbonX Editor**.
2. Abra o arquivo `.docm` no RibbonX Editor.
3. Insira o conteúdo do arquivo `src/CustomUI.xml`.
4. Salve e feche.

## Notas sobre UserForms

Os arquivos `.frm` contêm o código e a definição básica do formulário. Se houver problemas com a parte visual ao importar, você pode precisar ajustar o layout manualmente no Editor VBA, usando o código importado como base.
