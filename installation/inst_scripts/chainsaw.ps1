# =============================================================================
# CHAINSAW - Script Unificado de Gerenciamento
# =============================================================================
# Versão: 2.0.3
# Licença: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
# Compatibilidade: Windows 10+, PowerShell 5.1+
# Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
# =============================================================================

<#
.SYNOPSIS
    Script unificado para gerenciamento do CHAINSAW.

.DESCRIPTION
    Este script consolida todas as funcionalidades de gerenciamento:
    - install: Instala as configurações do Word
    - update-vba: Atualiza apenas o módulo VBA
    - export: Exporta configurações atuais
    - restore: Restaura backup anterior
    - enable-vba: Habilita acesso programático ao VBA
    
.PARAMETER Action
    Ação a ser executada: install, update-vba, export, restore, enable-vba

.PARAMETER Force
    Força a execução sem confirmação

.PARAMETER NoBackup
    Não cria backup (apenas para install)

.EXAMPLE
    .\chainsaw.ps1 install
    Instala o CHAINSAW

.EXAMPLE
    .\chainsaw.ps1 update-vba
    Atualiza apenas o módulo VBA

.EXAMPLE
    .\chainsaw.ps1 export
    Exporta configurações atuais

.EXAMPLE
    .\chainsaw.ps1 restore
    Restaura backup anterior

.EXAMPLE
    .\chainsaw.ps1 enable-vba
    Habilita acesso ao VBA
#>

[CmdletBinding()]
param(
    [Parameter(Position=0, Mandatory=$true)]
    [ValidateSet('install', 'update-vba', 'export', 'restore', 'enable-vba', 'disable-vba')]
    [string]$Action,
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$NoBackup,
    
    [Parameter()]
    [string]$SourcePath = ""
)

# Define o caminho base
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        $SourcePath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
}

# Redireciona para o script específico mantendo compatibilidade
switch ($Action) {
    'install' {
        $params = @{}
        if ($Force) { $params['Force'] = $true }
        if ($NoBackup) { $params['NoBackup'] = $true }
        if ($SourcePath) { $params['SourcePath'] = $SourcePath }
        
        & "$PSScriptRoot\install.ps1" @params
    }
    'update-vba' {
        $params = @{}
        if ($Force) { $params['Force'] = $true }
        
        & "$PSScriptRoot\update-vba-module.ps1" @params
    }
    'export' {
        $params = @{}
        if ($Force) { $params['Force'] = $true }
        
        & "$PSScriptRoot\export-config.ps1" @params
    }
    'restore' {
        $params = @{}
        if ($Force) { $params['Force'] = $true }
        
        & "$PSScriptRoot\restore-backup.ps1" @params
    }
    'enable-vba' {
        & "$PSScriptRoot\enable-vba-access.ps1"
    }
    'disable-vba' {
        & "$PSScriptRoot\enable-vba-access.ps1" -Disable
    }
}
