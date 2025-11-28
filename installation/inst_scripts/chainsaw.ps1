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
    Este script consolida as funcionalidades de gerenciamento:
    - install: Instala/atualiza as configurações do Word (inclui VBA, templates, customizações)
    - export: Exporta configurações atuais do Word
    
    O sistema de backup automático está integrado na instalação.
    Para restaurar um backup, renomeie a pasta chainsaw_backup para chainsaw.
    
.PARAMETER Action
    Ação a ser executada: install, export

.PARAMETER Force
    Força a execução sem confirmação

.PARAMETER NoBackup
    Não cria backup antes da instalação (não recomendado)

.PARAMETER SkipCustomizations
    Não importa customizações do Word durante instalação

.EXAMPLE
    .\chainsaw.ps1 install
    Instala/atualiza o CHAINSAW (cria backup automático)

.EXAMPLE
    .\chainsaw.ps1 install -Force
    Instala sem confirmação

.EXAMPLE
    .\chainsaw.ps1 export
    Exporta configurações atuais do Word
#>

[CmdletBinding()]
param(
    [Parameter(Position=0, Mandatory=$true)]
    [ValidateSet('install', 'export')]
    [string]$Action,
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$NoBackup,
    
    [Parameter()]
    [switch]$SkipCustomizations,
    
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

# Redireciona para o script específico
switch ($Action) {
    'install' {
        $params = @{}
        if ($Force) { $params['Force'] = $true }
        if ($NoBackup) { $params['NoBackup'] = $true }
        if ($SkipCustomizations) { $params['SkipCustomizations'] = $true }
        if ($SourcePath) { $params['SourcePath'] = $SourcePath }
        
        & "$PSScriptRoot\install.ps1" @params
    }
    'export' {
        $params = @{}
        if ($Force) { $params['Force'] = $true }
        
        & "$PSScriptRoot\export-config.ps1" @params
    }
}
