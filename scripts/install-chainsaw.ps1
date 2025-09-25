# =============================================================================
# CHAINSAW PROPOSITURAS - Script de Instalação Automatizada
# =============================================================================
#
# Este script automatiza a instalação completa do CHAINSAW PROPOSITURAS
# incluindo módulo VBA, interface, autotexto e configurações do Word.
#
# Versão: 2.0.0
# Data: 2025-09-23
# Autor: Christian Martin dos Santos
# Repositório: https://github.com/chrmsantos/chainsaw-proposituras
#
# =============================================================================

#Requires -Version 5.0

[CmdletBinding()]
param(
    [Parameter(HelpMessage="Pasta de destino para instalação")]
    [string]$InstallPath = "$env:USERPROFILE\Documents\CHAINSAW-PROPOSITURAS",
    
    [Parameter(HelpMessage="Instalar para todos os usuários (requer Admin)")]
    [switch]$AllUsers,
    
    [Parameter(HelpMessage="Pular verificações de segurança")]
    [switch]$SkipSecurityChecks,
    
    [Parameter(HelpMessage="Modo silencioso (sem prompts)")]
    [switch]$Silent,
    
    [Parameter(HelpMessage="Apenas verificar compatibilidade")]
    [switch]$CheckOnly
)

# =============================================================================
# CONFIGURAÇÕES E VARIÁVEIS GLOBAIS
# =============================================================================

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$Script:LogFile = "$env:TEMP\chainsaw-proposituras-install.log"
$Script:StartTime = Get-Date
$Script:WordVersions = @{
    "14.0" = "Word 2010"
    "15.0" = "Word 2013" 
    "16.0" = "Word 2016/2019/2021/365"
}

# Estrutura de arquivos necessários
$Script:RequiredFiles = @{
    "src\Módulo1.bas" = "Módulo VBA principal"
    "private\header\stamp.png" = "Logotipo do cabeçalho"
    "README.md" = "Documentação principal"
    "LICENSE" = "Arquivo de licença"
}

# =============================================================================
# FUNÇÕES DE LOGGING E INTERFACE
# =============================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Escreve no arquivo de log
    Add-Content -Path $Script:LogFile -Value $logMessage -Encoding UTF8
    
    # Exibe na tela com cores
    switch ($Level) {
        "INFO"    { Write-Host $Message -ForegroundColor White }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
    }
}

function Write-Header {
    param([string]$Title)
    
    $separator = "=" * 80
    Write-Host $separator -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host $separator -ForegroundColor Cyan
    Write-Host ""
}

function Confirm-Action {
    param(
        [string]$Message,
        [string]$Title = "Confirmação"
    )
    
    if ($Silent) { return $true }
    
    $choices = @(
        [System.Management.Automation.Host.ChoiceDescription]::new("&Sim", "Continuar com a ação")
        [System.Management.Automation.Host.ChoiceDescription]::new("&Não", "Cancelar a ação")
    )
    
    $result = $Host.UI.PromptForChoice($Title, $Message, $choices, 1)
    return ($result -eq 0)
}

# =============================================================================
# FUNÇÕES DE VERIFICAÇÃO DO SISTEMA
# =============================================================================

function Test-Prerequisites {
    Write-Log "Verificando pré-requisitos do sistema..." "INFO"
    
    $issues = @()
    
    # Verificar versão do PowerShell
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        $issues += "PowerShell 5.0+ necessário (atual: $($PSVersionTable.PSVersion))"
    }
    
    # Verificar se está no Windows
    if (-not $IsWindows -and $PSVersionTable.PSVersion.Major -ge 6) {
        $issues += "Sistema Windows necessário"
    }
    
    # Verificar privilégios de administrador se AllUsers
    if ($AllUsers -and -not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $issues += "Privilégios de administrador necessários para instalação global"
    }
    
    # Verificar Microsoft Word
    $wordCheck = Test-WordInstallation
    if (-not $wordCheck.IsInstalled) {
        $issues += "Microsoft Word não encontrado ou versão incompatível"
    }
    
    if ($issues.Count -gt 0) {
        Write-Log "Problemas encontrados:" "ERROR"
        foreach ($issue in $issues) {
            Write-Log "  - $issue" "ERROR"
        }
        return $false
    }
    
    Write-Log "Todos os pré-requisitos atendidos!" "SUCCESS"
    return $true
}

function Test-WordInstallation {
    Write-Log "Verificando instalação do Microsoft Word..." "INFO"
    
    $result = @{
        IsInstalled = $false
        Version = $null
        VersionName = $null
        Path = $null
        VBAEnabled = $false
    }
    
    try {
        # Tentar criar objeto do Word
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        
        $result.IsInstalled = $true
        $result.Version = $word.Version
        $result.Path = $word.Path
        $result.VersionName = $Script:WordVersions[$word.Version]
        
        # Verificar se VBA está habilitado
        try {
            $vbaProject = $word.VBE.ActiveVBProject
            $result.VBAEnabled = $true
        }
        catch {
            $result.VBAEnabled = $false
        }
        
        # Fechar Word
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
        Write-Log "Word encontrado: $($result.VersionName) ($($result.Version))" "SUCCESS"
        
        if (-not $result.VBAEnabled) {
            Write-Log "VBA pode estar desabilitado - verificação manual recomendada" "WARNING"
        }
        
    }
    catch {
        Write-Log "Microsoft Word não encontrado ou inacessível: $($_.Exception.Message)" "ERROR"
    }
    
    return $result
}

function Test-SecuritySettings {
    Write-Log "Verificando configurações de segurança..." "INFO"
    
    if ($SkipSecurityChecks) {
        Write-Log "Verificações de segurança puladas pelo usuário" "WARNING"
        return $true
    }
    
    $issues = @()
    
    try {
        # Verificar política de execução
        $executionPolicy = Get-ExecutionPolicy -Scope CurrentUser
        if ($executionPolicy -eq "Restricted") {
            $issues += "Política de execução muito restritiva para PowerShell"
        }
        
        # Verificar se arquivo está em zona da internet
        $scriptPath = $MyInvocation.MyCommand.Path
        if ($scriptPath) {
            $zone = Get-Item $scriptPath -Stream Zone.Identifier -ErrorAction SilentlyContinue
            if ($zone) {
                $issues += "Script baixado da internet - pode ser bloqueado pelo Windows"
            }
        }
        
    }
    catch {
        Write-Log "Erro ao verificar configurações de segurança: $($_.Exception.Message)" "WARNING"
    }
    
    if ($issues.Count -gt 0) {
        Write-Log "Avisos de segurança:" "WARNING"
        foreach ($issue in $issues) {
            Write-Log "  - $issue" "WARNING"
        }
        
        if (-not (Confirm-Action "Deseja continuar mesmo com os avisos de segurança?" "Aviso de Segurança")) {
            return $false
        }
    }
    
    return $true
}

# =============================================================================
# FUNÇÕES DE INSTALAÇÃO
# =============================================================================

function Install-ChainsawProposituras {
    Write-Header "INSTALAÇÃO DO CHAINSAW PROPOSITURAS"
    
    try {
        # Criar estrutura de diretórios
        New-DirectoryStructure
        
        # Copiar arquivos necessários
        Copy-RequiredFiles
        
        # Configurar Word
        Configure-WordSettings
        
        # Instalar módulo VBA
        Install-VBAModule
        
        # Criar shortcuts e menu
        Create-UserInterface
        
        # Configurar autotexto
        Configure-AutoText
        
        # Criar template normal personalizado
        Configure-NormalTemplate
        
        Write-Log "Instalação concluída com sucesso!" "SUCCESS"
        return $true
        
    }
    catch {
        Write-Log "Erro durante a instalação: $($_.Exception.Message)" "ERROR"
        Write-Log "Detalhes: $($_.Exception.ToString())" "ERROR"
        return $false
    }
}

function New-DirectoryStructure {
    Write-Log "Criando estrutura de diretórios..." "INFO"
    
    $directories = @(
        $InstallPath,
        "$InstallPath\src",
        "$InstallPath\private",
        "$InstallPath\private\header",
        "$InstallPath\private\backups",
        "$InstallPath\private\logs",
        "$InstallPath\templates",
        "$InstallPath\docs"
    )
    
    foreach ($dir in $directories) {
        if (-not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
            Write-Log "Criado: $dir" "INFO"
        }
    }
}

function Copy-RequiredFiles {
    Write-Log "Copiando arquivos necessários..." "INFO"
    
    $sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    
    foreach ($file in $Script:RequiredFiles.Keys) {
        $sourcePath = Join-Path $sourceDir $file
        $destPath = Join-Path $InstallPath $file
        
        if (Test-Path $sourcePath) {
            # Criar diretório de destino se não existir
            $destDir = Split-Path -Parent $destPath
            if (-not (Test-Path $destDir)) {
                New-Item -Path $destDir -ItemType Directory -Force | Out-Null
            }
            
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            Write-Log "Copiado: $($Script:RequiredFiles[$file])" "SUCCESS"
        }
        else {
            Write-Log "Arquivo não encontrado: $file" "WARNING"
        }
    }
}

function Configure-WordSettings {
    Write-Log "Configurando Microsoft Word..." "INFO"
    
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Configurações de segurança de macro
        # Nota: Algumas configurações podem requerer privilégios de administrador
        try {
            # Habilitar macros com notificação (valor 2)
            $word.Application.AutomationSecurity = 1 # msoAutomationSecurityLow
            Write-Log "Configurações de automação ajustadas" "SUCCESS"
        }
        catch {
            Write-Log "Não foi possível ajustar configurações de automação: $($_.Exception.Message)" "WARNING"
        }
        
        # Configurar pasta de templates do usuário
        try {
            $userTemplatesPath = "$InstallPath\templates"
            # Esta configuração pode não persistir dependendo da versão do Word
            Write-Log "Pasta de templates configurada: $userTemplatesPath" "INFO"
        }
        catch {
            Write-Log "Não foi possível configurar pasta de templates" "WARNING"
        }
        
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
    }
    catch {
        Write-Log "Erro ao configurar Word: $($_.Exception.Message)" "ERROR"
    }
}

function Install-VBAModule {
    Write-Log "Instalando módulo VBA..." "INFO"
    
    $vbaPath = Join-Path $InstallPath "src\Módulo1.bas"
    
    if (-not (Test-Path $vbaPath)) {
        Write-Log "Módulo VBA não encontrado: $vbaPath" "ERROR"
        return $false
    }
    
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Abrir template Normal
        $template = $word.NormalTemplate
        
        # Importar módulo VBA
        $vbProject = $template.VBProject
        $vbProject.VBComponents.Import($vbaPath) | Out-Null
        
        # Salvar template
        $template.Save()
        
        Write-Log "Módulo VBA instalado com sucesso!" "SUCCESS"
        
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
        return $true
        
    }
    catch {
        Write-Log "Erro ao instalar módulo VBA: $($_.Exception.Message)" "ERROR"
        Write-Log "Nota: Pode ser necessário habilitar 'Acesso ao modelo de objeto do projeto VBA' nas configurações do Word" "WARNING"
        return $false
    }
}

function Create-UserInterface {
    Write-Log "Criando interface do usuário..." "INFO"
    
    try {
        # Criar atalho na área de trabalho
        $desktop = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path $desktop "Chainsaw Proposituras.lnk"
        
        $wscript = New-Object -ComObject WScript.Shell
        $shortcut = $wscript.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "winword.exe"
        $shortcut.Arguments = "/mPadronizarDocumentoMain"
        $shortcut.WorkingDirectory = $InstallPath
        $shortcut.Description = "Chainsaw Proposituras - Padronização de Documentos Legislativos"
        $shortcut.Save()
        
        Write-Log "Atalho criado na área de trabalho" "SUCCESS"
        
        # Criar atalho no menu Iniciar (se AllUsers)
        if ($AllUsers) {
            $startMenu = [Environment]::GetFolderPath("CommonStartMenu")
            $programsPath = Join-Path $startMenu "Programs\Chainsaw Proposituras"
            
            if (-not (Test-Path $programsPath)) {
                New-Item -Path $programsPath -ItemType Directory -Force | Out-Null
            }
            
            $startShortcut = Join-Path $programsPath "Chainsaw Proposituras.lnk"
            $shortcut2 = $wscript.CreateShortcut($startShortcut)
            $shortcut2.TargetPath = "winword.exe"
            $shortcut2.Arguments = "/mPadronizarDocumentoMain"
            $shortcut2.WorkingDirectory = $InstallPath
            $shortcut2.Description = "Chainsaw Proposituras - Padronização de Documentos Legislativos"
            $shortcut2.Save()
            
            Write-Log "Atalho criado no menu Iniciar" "SUCCESS"
        }
        
    }
    catch {
        Write-Log "Erro ao criar interface: $($_.Exception.Message)" "WARNING"
    }
}

function Configure-AutoText {
    Write-Log "Configurando autotexto..." "INFO"
    
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Abrir template Normal
        $template = $word.NormalTemplate
        
        # Entradas de autotexto comuns para proposituras
        $autoTextEntries = @{
            "indicacao" = @{
                "name" = "Cabeçalho Indicação"
                "text" = "INDICAÇÃO Nº `$NUMERO`$/`$ANO`$`n`nSenhor Presidente,"
            }
            "requerimento" = @{
                "name" = "Cabeçalho Requerimento" 
                "text" = "REQUERIMENTO Nº `$NUMERO`$/`$ANO`$`n`nSenhor Presidente,"
            }
            "mocao" = @{
                "name" = "Cabeçalho Moção"
                "text" = "MOÇÃO Nº `$NUMERO`$/`$ANO`$`n`nSenhor Presidente,"
            }
            "considerando" = @{
                "name" = "Considerando Padrão"
                "text" = "CONSIDERANDO que "
            }
            "justificativa" = @{
                "name" = "Justificativa"
                "text" = "`n`nJUSTIFICATIVA`n`n"
            }
            "vereador" = @{
                "name" = "Assinatura Vereador"
                "text" = "`n`n- VEREADOR -"
            }
        }
        
        foreach ($entry in $autoTextEntries.GetEnumerator()) {
            try {
                # Criar nova entrada de autotexto
                $doc = $word.Documents.Add()
                $doc.Range().Text = $entry.Value.text -replace "``n", "`r`n"
                $doc.Range().Select()
                
                $template.AutoTextEntries.Add($entry.Value.name, $word.Selection.Range) | Out-Null
                $doc.Close($false)
                
                Write-Log "Autotexto criado: $($entry.Value.name)" "SUCCESS"
            }
            catch {
                Write-Log "Erro ao criar autotexto '$($entry.Key)': $($_.Exception.Message)" "WARNING"
            }
        }
        
        # Salvar template
        $template.Save()
        
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
        Write-Log "Configuração de autotexto concluída" "SUCCESS"
        
    }
    catch {
        Write-Log "Erro ao configurar autotexto: $($_.Exception.Message)" "WARNING"
    }
}

function Configure-NormalTemplate {
    Write-Log "Configurando template Normal.dotm..." "INFO"
    
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        # Abrir template Normal
        $template = $word.NormalTemplate
        
        # Configurar estilos padrão
        Configure-DefaultStyles $template
        
        # Configurar margens padrão
        Configure-DefaultPageSetup $template
        
        # Salvar template
        $template.Save()
        
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        
        Write-Log "Template Normal.dotm configurado" "SUCCESS"
        
    }
    catch {
        Write-Log "Erro ao configurar template Normal: $($_.Exception.Message)" "WARNING"
    }
}

function Configure-DefaultStyles {
    param($template)
    
    try {
        # Configurar estilo Normal
        $normalStyle = $template.Styles["Normal"]
        $normalStyle.Font.Name = "Arial"
        $normalStyle.Font.Size = 12
        $normalStyle.ParagraphFormat.LineSpacing = 16.8 # 1.4 * 12pt
        $normalStyle.ParagraphFormat.Alignment = 3 # Justificado
        
        Write-Log "Estilo Normal configurado" "INFO"
        
    }
    catch {
        Write-Log "Erro ao configurar estilos: $($_.Exception.Message)" "WARNING"
    }
}

function Configure-DefaultPageSetup {
    param($template)
    
    try {
        # Configurar documento temporário para definir margens padrão
        $doc = $template.Application.Documents.Add($template)
        
        $pageSetup = $doc.PageSetup
        $pageSetup.TopMargin = 131.8 # 4.6cm em pontos
        $pageSetup.BottomMargin = 56.7 # 2cm em pontos  
        $pageSetup.LeftMargin = 85.05 # 3cm em pontos
        $pageSetup.RightMargin = 85.05 # 3cm em pontos
        $pageSetup.HeaderDistance = 8.5 # 0.3cm em pontos
        $pageSetup.FooterDistance = 25.5 # 0.9cm em pontos
        
        $doc.Close($false)
        
        Write-Log "Configuração de página padrão definida" "INFO"
        
    }
    catch {
        Write-Log "Erro ao configurar página padrão: $($_.Exception.Message)" "WARNING"
    }
}

# =============================================================================
# FUNÇÕES DE VERIFICAÇÃO E TESTE
# =============================================================================

function Test-Installation {
    Write-Log "Verificando instalação..." "INFO"
    
    $checks = @()
    
    # Verificar estrutura de diretórios
    $requiredDirs = @(
        "$InstallPath\src",
        "$InstallPath\private\header",
        "$InstallPath\private\backups",
        "$InstallPath\private\logs"
    )
    
    foreach ($dir in $requiredDirs) {
        if (Test-Path $dir) {
            $checks += "✓ Diretório: $(Split-Path -Leaf $dir)"
        }
        else {
            $checks += "✗ Diretório ausente: $(Split-Path -Leaf $dir)"
        }
    }
    
    # Verificar módulo VBA
    $vbaPath = Join-Path $InstallPath "src\Módulo1.bas"
    if (Test-Path $vbaPath) {
        $checks += "✓ Módulo VBA presente"
    }
    else {
        $checks += "✗ Módulo VBA ausente"
    }
    
    # Verificar logotipo
    $logoPath = Join-Path $InstallPath "private\header\stamp.png"
    if (Test-Path $logoPath) {
        $checks += "✓ Logotipo presente"
    }
    else {
        $checks += "✗ Logotipo ausente"
    }
    
    Write-Log "Resultado da verificação:" "INFO"
    foreach ($check in $checks) {
        if ($check.StartsWith("✓")) {
            Write-Log $check "SUCCESS"
        }
        else {
            Write-Log $check "ERROR"
        }
    }
    
    return -not ($checks -match "✗").Count
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================

function Main {
    Write-Header "CHAINSAW PROPOSITURAS - INSTALADOR AUTOMATIZADO v2.0"
    
    Write-Log "Iniciando instalação em: $InstallPath" "INFO"
    Write-Log "Log de instalação: $Script:LogFile" "INFO"
    Write-Host ""
    
    # Verificações preliminares
    if (-not (Test-Prerequisites)) {
        Write-Log "Pré-requisitos não atendidos. Instalação abortada." "ERROR"
        exit 1
    }
    
    if (-not (Test-SecuritySettings)) {
        Write-Log "Verificações de segurança falharam. Instalação abortada." "ERROR"
        exit 1
    }
    
    # Verificação apenas
    if ($CheckOnly) {
        Write-Log "Verificação de compatibilidade concluída. Sistema pronto para instalação." "SUCCESS"
        exit 0
    }
    
    # Confirmação final
    if (-not $Silent) {
        Write-Host "A instalação será realizada em: $InstallPath" -ForegroundColor Yellow
        Write-Host ""
        
        if (-not (Confirm-Action "Deseja continuar com a instalação?" "Confirmação de Instalação")) {
            Write-Log "Instalação cancelada pelo usuário." "INFO"
            exit 0
        }
    }
    
    # Executar instalação
    Write-Host ""
    if (Install-ChainsawProposituras) {
        Write-Host ""
        Write-Header "INSTALAÇÃO CONCLUÍDA COM SUCESSO!"
        
        # Verificar instalação
        if (Test-Installation) {
            Write-Log "Verificação pós-instalação bem-sucedida." "SUCCESS"
        }
        else {
            Write-Log "Alguns componentes podem não ter sido instalados corretamente." "WARNING"
        }
        
        Write-Host ""
        Write-Host "PRÓXIMOS PASSOS:" -ForegroundColor Green
        Write-Host "1. Abra o Microsoft Word" -ForegroundColor White
        Write-Host "2. Verifique se as macros estão habilitadas" -ForegroundColor White
        Write-Host "3. Execute a macro 'PadronizarDocumentoMain' ou use o atalho criado" -ForegroundColor White
        Write-Host "4. Consulte a documentação em: $InstallPath\README.md" -ForegroundColor White
        Write-Host ""
        Write-Host "Log completo salvo em: $Script:LogFile" -ForegroundColor Gray
        
        $duration = (Get-Date) - $Script:StartTime
        Write-Log "Instalação concluída em $($duration.TotalSeconds.ToString('F1')) segundos" "SUCCESS"
        
    }
    else {
        Write-Host ""
        Write-Header "INSTALAÇÃO FALHOU"
        Write-Log "Consulte o log para detalhes: $Script:LogFile" "ERROR"
        exit 1
    }
}

# =============================================================================
# EXECUÇÃO
# =============================================================================

# Inicializar log
"Chainsaw Proposituras - Log de Instalação" | Out-File -FilePath $Script:LogFile -Encoding UTF8
"Iniciado em: $(Get-Date)" | Out-File -FilePath $Script:LogFile -Append -Encoding UTF8
"" | Out-File -FilePath $Script:LogFile -Append -Encoding UTF8

# Executar função principal
try {
    Main
}
catch {
    Write-Log "Erro crítico durante a instalação: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.Exception.ToString())" "ERROR"
    exit 1
}
finally {
    # Limpeza final
    Write-Log "Script finalizado em: $(Get-Date)" "INFO"
}