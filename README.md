# Projeto de Automação APA 7 para Word

Este projeto contém macros VBA para automação de citações e referências no formato APA 7ª Edição.

## Funcionalidades

- **Inserir Citação**: Adiciona citações narrativas ou parentéticas.
- **Gerenciar Referências**: Adiciona, edita e remove referências bibliográficas.
- **Gerar Bibliografia**: Cria a lista de referências ao final do documento.
- **Validação**: Verifica se todas as referências possuem os campos obrigatórios.

## Estrutura do Projeto

- **src/**: Contém todo o código fonte.
  - `.cls` - Classes VBA (Autor, Referência, etc).
  - `.bas` - Módulos Padrão (Lógica principal).
  - `.frm` - Formulários de Interface (UserForms).
  - `CustomUI.xml` - Definição da barra de ferramentas (Ribbon).
- **build.ps1**: Script PowerShell para gerar o arquivo final `.docm` automaticamente.

## Como Instalar/Compilar

### Método Automático (Recomendado)

1. Abra o PowerShell nesta pasta.
2. Execute o script de build:

    ```powershell
    .\build.ps1
    ```

3. O arquivo **FormatadorAPA7.docm** será criado na raiz.
4. **Pronto!** A barra de ferramentas "APA 7" já estará disponível ao abrir o arquivo.

> **Nota:** Certifique-se de que a opção *"Confiar no acesso ao modelo de objeto do projeto do VBA"* esteja habilitada nas configurações de Macro do Word (Central de Confiabilidade).

### Instalação Manual (Opcional)

Se preferir fazer tudo manualmente:

1. Crie um novo arquivo Word Habilitado para Macro (`.docm`).
2. Abra o editor VBA (`Alt+F11`).
3. Importe todos os arquivos da pasta `src/`.
4. Adicione as referências "Microsoft Scripting Runtime" e "Microsoft XML, v6.0".
5. Use o **Office RibbonX Editor** para injetar o `src/CustomUI.xml` no arquivo.
