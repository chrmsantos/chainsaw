[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent (Get-Item -LiteralPath $MyInvocation.MyCommand.Path).FullName
}
$installScript = Join-Path $scriptRoot 'install.ps1'
$installationFolder = Split-Path -Parent $scriptRoot
$logDirectory = Join-Path $installationFolder 'inst_docs\inst_logs'

function Show-Header {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host '  CHAINSAW - Instalador' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkGray
    Write-Host ''
}

function Ask-YesNo {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [Parameter()] [bool]$DefaultYes = $true
    )

    $suffix = if ($DefaultYes) { '[S/n]' } else { '[s/N]' }
    while ($true) {
        $answer = Read-Host "$Message $suffix"
        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $DefaultYes
        }

        switch ($answer.Trim().ToLowerInvariant()) {
            's' { return $true }
            'sim' { return $true }
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'nao' { return $false }
            'no' { return $false }
            default { Write-Host 'Digite apenas S ou N.' -ForegroundColor Yellow }
        }
    }
}

function Read-MenuOption {
    $options = @(
        'Instalacao padrao (recomendada)',
        'Instalacao personalizada',
        'Cancelar'
    )

    Write-Host 'Escolha uma opcao:' -ForegroundColor White
    for ($i = 0; $i -lt $options.Count; $i++) {
        $index = $i + 1
        Write-Host "  $index) $($options[$i])" -ForegroundColor Gray
    }

    while ($true) {
        $inputValue = Read-Host 'Opcao'
        if ([int]::TryParse($inputValue, [ref]$null)) {
            $number = [int]$inputValue
            if ($number -ge 1 -and $number -le $options.Count) {
                return $number
            }
        }
        Write-Host 'Escolha um numero valido.' -ForegroundColor Yellow
    }
}

function Get-LatestLogPath {
    if (-not (Test-Path $logDirectory)) {
        return $null
    }

    $file = Get-ChildItem -Path $logDirectory -Filter 'install_*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    return $file?.FullName
}

function Start-Installer {
    param(
        [Parameter()] [bool]$Force,
        [Parameter()] [bool]$SkipBackup,
        [Parameter()] [bool]$SkipCustomizations
    )

    $arguments = @('-ExecutionPolicy','Bypass','-NoProfile','-NoLogo','-File',"`"$installScript`"")
    if ($Force) { $arguments += '-Force' }
    if ($SkipBackup) { $arguments += '-NoBackup' }
    if ($SkipCustomizations) { $arguments += '-SkipCustomizations' }

    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -Wait -PassThru
    return $process.ExitCode
}

try {
    if (-not (Test-Path $installScript)) {
        Write-Host '[ERRO] install.ps1 nao encontrado.' -ForegroundColor Red
        exit 1
    }

    Show-Header
    $choice = Read-MenuOption

    if ($choice -eq 3) {
        Write-Host ''
        Write-Host 'Operacao cancelada pelo usuario.' -ForegroundColor Yellow
        exit 0
    }

    $forceInstall = $true
    $skipBackup = $false
    $skipCustom = $false

    if ($choice -eq 2) {
        $forceInstall = Ask-YesNo 'Deseja executar sem prompts internos do instalador?' $false
        $skipBackup = Ask-YesNo 'Deseja pular o backup automatico? (nao recomendado)' $false
        $skipCustom = Ask-YesNo 'Deseja ignorar a importacao de customizacoes?' $false
    }

    $proceed = Ask-YesNo 'Deseja iniciar a instalacao agora?'
    if (-not $proceed) {
        Write-Host ''
        Write-Host 'Instalacao cancelada antes da execucao.' -ForegroundColor Yellow
        exit 0
    }

    Write-Host ''
    Write-Host 'Executando instalador...' -ForegroundColor Cyan
    $exitCode = Start-Installer -Force:$forceInstall -SkipBackup:$skipBackup -SkipCustomizations:$skipCustom
    $logPath = Get-LatestLogPath

    if ($exitCode -eq 0) {
        Write-Host ''
        Write-Host 'Instalacao concluida com sucesso.' -ForegroundColor Green
        if ($logPath) {
            Write-Host "Log: $logPath" -ForegroundColor DarkGray
        }
        exit 0
    }

    Write-Host ''
    Write-Host 'Um erro foi identificado durante a instalacao.' -ForegroundColor Red
    if ($logPath) {
        Write-Host "Consulte o log para detalhes: $logPath" -ForegroundColor Yellow
    }
    else {
        Write-Host 'Consulte os logs em installation\\inst_docs\\inst_logs para detalhes.' -ForegroundColor Yellow
    }
    exit $exitCode
}
catch {
    Write-Host ''
    Write-Host 'O instalador encontrou uma falha inesperada.' -ForegroundColor Red
    $logPath = Get-LatestLogPath
    if ($logPath) {
        Write-Host "Consulte o log para detalhes: $logPath" -ForegroundColor Yellow
    }
    else {
        Write-Host 'Consulte os logs em installation\\inst_docs\\inst_logs para detalhes.' -ForegroundColor Yellow
    }
    exit 1
}
