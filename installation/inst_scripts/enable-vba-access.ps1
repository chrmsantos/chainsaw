# =============================================================================
# CHAINSAW - Habilitar Acesso Programatico ao VBA
# =============================================================================
# Versao: 1.0.0
# Licenca: GNU GPLv3
# Autor: Christian Martin dos Santos
# =============================================================================

<#
.SYNOPSIS
    Habilita o acesso programatico ao modelo de objeto do projeto VBA.

.DESCRIPTION
    Este script configura o Word para permitir acesso programatico ao VBA,
    necessario para exportar e importar modulos VBA automaticamente.
    
    Modifica a chave do registro:
    HKCU:\Software\Microsoft\Office\16.0\Word\Security
    AccessVBOM = 1
    
.NOTES
    - Nao requer privilegios de administrador (modifica apenas HKCU)
    - Funciona com Word 2016/2019/2021/365
    - Reversivel executando com -Disable

.EXAMPLE
    .\enable-vba-access.ps1
    Habilita o acesso ao VBA

.EXAMPLE
    .\enable-vba-access.ps1 -Disable
    Desabilita o acesso ao VBA (restaura configuracao padrao)
#>

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$Disable
)

$ErrorActionPreference = "Stop"

# =============================================================================
# CONFIGURACOES
# =============================================================================

$ColorSuccess = "Green"
$ColorWarning = "Yellow"
$ColorError = "Red"
$ColorInfo = "Cyan"

# Versoes do Office suportadas
$OfficeVersions = @(
    "16.0",  # Office 2016/2019/2021/365
    "15.0",  # Office 2013
    "14.0"   # Office 2010
)

# =============================================================================
# FUNCOES
# =============================================================================

function Write-ColorHost {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Get-WordVersion {
    <#
    .SYNOPSIS
        Detecta a versao do Word instalada.
    #>
    foreach ($version in $OfficeVersions) {
        $regPath = "HKCU:\Software\Microsoft\Office\$version\Word"
        if (Test-Path $regPath) {
            return $version
        }
    }
    return $null
}

function Test-WordRunning {
    <#
    .SYNOPSIS
        Verifica se o Word esta em execucao.
    #>
    $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Enable-VBAAccess {
    <#
    .SYNOPSIS
        Habilita o acesso programatico ao VBA.
    #>
    param(
        [string]$WordVersion
    )
    
    $regPath = "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security"
    
    try {
        # Cria a chave Security se nao existir
        if (-not (Test-Path $regPath)) {
            Write-ColorHost "Criando chave de registro: $regPath" -Color $ColorInfo
            New-Item -Path $regPath -Force | Out-Null
        }
        
        # Define AccessVBOM = 1
        Write-ColorHost "Habilitando acesso ao VBA Object Model..." -Color $ColorInfo
        Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -Force
        
        # Verifica se foi aplicado
        $currentValue = Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
        
        if ($currentValue.AccessVBOM -eq 1) {
            Write-ColorHost "[OK] Acesso ao VBA habilitado com sucesso!" -Color $ColorSuccess
            return $true
        }
        else {
            Write-ColorHost "[ERRO] Falha ao verificar configuracao" -Color $ColorError
            return $false
        }
    }
    catch {
        Write-ColorHost "[ERRO] Falha ao modificar registro: $_" -Color $ColorError
        return $false
    }
}

function Disable-VBAAccess {
    <#
    .SYNOPSIS
        Desabilita o acesso programatico ao VBA (restaura padrao).
    #>
    param(
        [string]$WordVersion
    )
    
    $regPath = "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security"
    
    try {
        if (Test-Path $regPath) {
            Write-ColorHost "Desabilitando acesso ao VBA Object Model..." -Color $ColorInfo
            
            # Remove a chave ou define como 0
            Remove-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
            
            Write-ColorHost "[OK] Acesso ao VBA desabilitado (restaurado ao padrao)" -Color $ColorSuccess
            return $true
        }
        else {
            Write-ColorHost "[INFO] Chave de registro nao existe - nada a fazer" -Color $ColorInfo
            return $true
        }
    }
    catch {
        Write-ColorHost "[ERRO] Falha ao modificar registro: $_" -Color $ColorError
        return $false
    }
}

function Show-VBAAccessStatus {
    <#
    .SYNOPSIS
        Mostra o status atual do acesso ao VBA.
    #>
    param(
        [string]$WordVersion
    )
    
    $regPath = "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security"
    
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor DarkGray
    Write-Host "  Status do Acesso ao VBA Object Model" -ForegroundColor White
    Write-Host "================================================================" -ForegroundColor DarkGray
    Write-Host ""
    
    if (Test-Path $regPath) {
        $accessVBOM = Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
        
        if ($null -ne $accessVBOM -and $accessVBOM.AccessVBOM -eq 1) {
            Write-ColorHost "  Status: [HABILITADO]" -Color $ColorSuccess
            Write-ColorHost "  Valor: AccessVBOM = 1" -Color $ColorInfo
            Write-Host "  Caminho: $regPath" -ForegroundColor Gray
        }
        else {
            Write-ColorHost "  Status: [DESABILITADO]" -Color $ColorWarning
            Write-ColorHost "  Valor: AccessVBOM nao definido ou = 0" -Color $ColorInfo
            Write-Host "  Caminho: $regPath" -ForegroundColor Gray
        }
    }
    else {
        Write-ColorHost "  Status: [DESABILITADO]" -Color $ColorWarning
        Write-ColorHost "  Valor: Chave de registro nao existe" -Color $ColorInfo
        Write-Host "  Caminho esperado: $regPath" -ForegroundColor Gray
    }
    
    Write-Host ""
}

# =============================================================================
# EXECUCAO PRINCIPAL
# =============================================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  CHAINSAW - Configuracao de Acesso ao VBA Object Model" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Detecta versao do Word
Write-ColorHost "Detectando versao do Word instalada..." -Color $ColorInfo
$wordVersion = Get-WordVersion

if ($null -eq $wordVersion) {
    Write-Host ""
    Write-ColorHost "[ERRO] Microsoft Word nao encontrado!" -Color $ColorError
    Write-Host ""
    Write-Host "  O Word nao parece estar instalado neste computador." -ForegroundColor Gray
    Write-Host "  Versoes suportadas: Word 2010/2013/2016/2019/2021/365" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Pressione qualquer tecla para sair..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-ColorHost "[OK] Word $wordVersion detectado" -Color $ColorSuccess
Write-Host ""

# Verifica se Word esta em execucao
if (Test-WordRunning) {
    Write-Host "================================================================" -ForegroundColor Yellow
    Write-Host "                    [AVISO] IMPORTANTE" -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-ColorHost "O Microsoft Word esta em execucao!" -Color $ColorWarning
    Write-Host ""
    Write-Host "  As alteracoes serao aplicadas, mas o Word precisa ser" -ForegroundColor Gray
    Write-Host "  REINICIADO para que as configuracoes tenham efeito." -ForegroundColor Gray
    Write-Host ""
    
    $continue = Read-Host "Deseja continuar mesmo assim? (S/N)"
    if ($continue -notmatch '^[Ss]$') {
        Write-ColorHost "[INFO] Operacao cancelada pelo usuario" -Color $ColorInfo
        exit 0
    }
    Write-Host ""
}

# Mostra status atual
Show-VBAAccessStatus -WordVersion $wordVersion

# Executa operacao
if ($Disable) {
    Write-Host "Desabilitando acesso ao VBA..." -ForegroundColor Yellow
    Write-Host ""
    
    if (Disable-VBAAccess -WordVersion $wordVersion) {
        Write-Host ""
        Show-VBAAccessStatus -WordVersion $wordVersion
        
        Write-Host "================================================================" -ForegroundColor Green
        Write-Host "              CONFIGURACAO RESTAURADA" -ForegroundColor Green
        Write-Host "================================================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "  O acesso programatico ao VBA foi desabilitado." -ForegroundColor Gray
        Write-Host "  Configuracao de seguranca padrao do Word restaurada." -ForegroundColor Gray
        Write-Host ""
    }
    else {
        Write-Host ""
        Write-Host "Pressione qualquer tecla para sair..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}
else {
    Write-Host "Habilitando acesso ao VBA..." -ForegroundColor Cyan
    Write-Host ""
    
    if (Enable-VBAAccess -WordVersion $wordVersion) {
        Write-Host ""
        Show-VBAAccessStatus -WordVersion $wordVersion
        
        Write-Host "================================================================" -ForegroundColor Green
        Write-Host "           CONFIGURACAO APLICADA COM SUCESSO!" -ForegroundColor Green
        Write-Host "================================================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "  O Word agora permite acesso programatico ao VBA." -ForegroundColor Gray
        Write-Host "  Voce pode executar export-config.ps1 sem erros." -ForegroundColor Gray
        Write-Host ""
        
        if (Test-WordRunning) {
            Write-ColorHost "  [LEMBRE-SE] Reinicie o Word para que as mudancas tenham efeito!" -Color $ColorWarning
            Write-Host ""
        }
        
        Write-Host "  Para desabilitar novamente, execute:" -ForegroundColor DarkGray
        Write-Host "    .\enable-vba-access.ps1 -Disable" -ForegroundColor DarkGray
        Write-Host ""
    }
    else {
        Write-Host ""
        Write-Host "Pressione qualquer tecla para sair..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}

Write-Host "Pressione qualquer tecla para sair..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
