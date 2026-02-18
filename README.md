# Projeto de Automação APA 7 para Word

Este projeto contém macros VBA para autmação de citações e referências no formato APA 7ª Edição.

## Estrutura do Projeto

* **src/**: Contém todo o código fonte.
  * `.cls` - Classes VBA (Autor, Referência, etc).
  * `.bas` - Módulos Padrão (Lógica principal).
  * `.frm` - Formulários de Interface (UserForms).
  * `CustomUI.xml` - Definição da barra de ferramentas (Ribbon).
* **build.ps1**: Script PowerShell para gerar o arquivo final `.docm` automaticamente.

## Como Instalar/Compilar

### Método Automático (Recomendado)

1. Abra o PowerShell nesta pasta.
2. Execute o script de build:

    ```powershell
    .\build.ps1
    ```

3. O arquivo **FormatadorAPA7.docm** será criado na raiz.

> **Nota:** Certifique-se de que a opção *"Confiar no acesso ao modelo de objeto do projeto do VBA"* esteja habilitada nas configurações de Macro do Word (Central de Confiabilidade).

### Método Manual

1. Crie um novo arquivo Word Habilitado para Macro (`.docm`).
2. Abra o editor VBA (`Alt+F11`).
3. Importe todos os arquivos da pasta `src/`.
4. Adicione as referências "Microsoft Scripting Runtime" e "Microsoft XML, v6.0".

## Customização da Ribbon (Interface)

O script `build.ps1` importa o código VBA, mas a interface visual (botões no topo do Word) definida em `src/CustomUI.xml` precisa ser injetada separadamente se o script não o fizer (o script atual foca no VBA).

Para ver a aba "APA 7":

1. Baixe o **Office RibbonX Editor**.
2. Abra o arquivo `FormatadorAPA7.docm`.
3. Cole o conteúdo de `src/CustomUI.xml`.
4. Salve.
