<#
.SYNOPSIS
    Compilador automático para o projeto VBA APA 7.0
    Gera o arquivo 'FormatadorAPA7.docm' a partir dos fontes em 'src/'.
.DESCRIPTION
    Este script cria um novo documento Word, habilita macros, adiciona
    as referências necessárias (VBA, Scripting, XML) e importa
    os arquivos do diretório de origem.
.NOTES
    Requisitos:
    - Microsoft Word instalado.
    - "Trust access to the VBA project object model" habilitado no Word
      (Opções > Central de Confiabilidade > Configurações de Macro).
#>

$sourceDir = "$PSScriptRoot\src"
$outputFile = "$PSScriptRoot\FormatadorAPA7.docm"
$word = New-Object -ComObject Word.Application
$word.Visible = $true # Tornar visível para debug visual, se necessário

try {
    Write-Host "Criando novo documento Word..." -ForegroundColor Cyan
    $doc = $word.Documents.Add()
    
    # Salvar primeiro como .docm para habilitar o projeto VBA
    Write-Host "Salvando como formato Macro-Enabled (.docm)..." -ForegroundColor Cyan
    $doc.SaveAs([ref]$outputFile, [ref]13) # wdFormatXMLDocumentMacroEnabled = 13
    
    $vbProject = $doc.VBProject
    
    # ---------------------------------------------------------
    # 1. Adicionar Referências Externas
    # ---------------------------------------------------------
    Write-Host "Adicionando Referências..." -ForegroundColor Green
    
    # Microsoft Scripting Runtime (scrrun.dll)
    try {
        $vbProject.References.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
        Write-Host "  [OK] Microsoft Scripting Runtime" -ForegroundColor Green
    } catch {
        Write-Warning "  [FALHA] Microsoft Scripting Runtime já existe ou erro: $_"
    }

    # Microsoft XML, v6.0 (msxml6.dll)
    try {
        $vbProject.References.AddFromGuid("{F5078F18-C551-11D3-89B9-0000F81FE221}", 6, 0)
        Write-Host "  [OK] Microsoft XML, v6.0" -ForegroundColor Green
    } catch {
        Write-Warning "  [FALHA] Microsoft XML, v6.0 já existe ou erro: $_"
    }

    # ---------------------------------------------------------
    # 2. Importar Arquivos Fonte
    # ---------------------------------------------------------
    Write-Host "Importando módulos VBA..." -ForegroundColor Yellow
    
    $files = Get-ChildItem -Path $sourceDir -Include *.cls, *.bas, *.frm -Recurse
    
    foreach ($file in $files) {
        Write-Host "  Importando: $($file.Name)"
        try {
            $vbProject.VBComponents.Import($file.FullName)
        } catch {
            Write-Error "  Erro ao importar $($file.Name): $_"
        }
    }
    
    # ---------------------------------------------------------
    # 3. Remover Módulo Padrão Vazio (Opcional)
    # ---------------------------------------------------------
    # O Word cria 'Module1' por padrão. Podemos remover se desejado.
    # try {
    #     $component = $vbProject.VBComponents.Item("Module1")
    #     $vbProject.VBComponents.Remove($component)
    # } catch {}

    # Salvar e Fechar
    $doc.Save()
    Write-Host "----- SUCESSO! -----" -ForegroundColor Cyan
    Write-Host "Arquivo gerado em: $outputFile" -ForegroundColor Cyan
    
} catch {
    Write-Error "ERRO CRÍTICO: $_"
    Write-Warning "Certifique-se de que 'Confiar no acesso ao modelo de objeto do projeto do VBA' está habilitado no Word."
} finally {
    # Fechar documento e Word
    if ($doc) { $doc.Close() }
    $word.Quit()
    
    # Limpar objetos COM da memória
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    Remove-Variable word, doc
}
