# =============================================================================
# CHAINSAW - Script de Instala��o de Configura��es do Word
# =============================================================================
# Vers�o: 2.0.0
# Licen�a: GNU GPLv3 (https://www.gnu.org/licenses/gpl-3.0.html)
# Compatibilidade: Windows 10+, PowerShell 5.1+
# Autor: Christian Martin dos Santos (chrmsantos@protonmail.com)
# =============================================================================

<#
.SYNOPSIS
 [X]  [X] Instala as configura��es do Word do sistema CHAINSAW para o usu�rio atual.

.DESCRIPTION
 [X]  [X] Este script realiza as seguintes opera��es:
 [X]  [X] 1. Copia o arquivo stamp.png para a pasta do usu�rio
 [X]  [X] 2. Faz backup da pasta Templates atual
 [X]  [X] 3. Copia os novos Templates
 [X]  [X] 4. Detecta e importa personaliza��es do Word (se encontradas)
 [X]  [X] 5. Registra todas as opera��es em arquivo de log
 [X]  [X] 
 [X]  [X] Se uma pasta 'exported-config' for encontrada no diret�rio do script,
 [X]  [X] as personaliza��es do Word (Ribbon, Partes R�pidas, etc.) ser�o 
 [X]  [X] automaticamente importadas.

.PARAMETER SourcePath
 [X]  [X] Caminho base dos arquivos. Padr�o: pasta onde o script est� localizado

.PARAMETER Force
 [X]  [X] For�a a instala��o sem confirma��o do usu�rio.

.PARAMETER NoBackup
 [X]  [X] N�o cria backup da pasta Templates existente (n�o recomendado).

.PARAMETER SkipCustomizations
 [X]  [X] N�o importa personaliza��es do Word mesmo se encontradas.

.EXAMPLE
 [X]  [X] .\install.ps1
 [X]  [X] Executa a instala��o com confirma��o do usu�rio.

.EXAMPLE
 [X]  [X] .\install.ps1 -Force
 [X]  [X] Executa a instala��o sem confirma��o.

.EXAMPLE
 [X]  [X] .\install.ps1 -SkipCustomizations
 [X]  [X] Instala apenas Templates, sem importar personaliza��es.

.NOTES
 [X]  [X] Requer permiss�es de escrita nas pastas do usu�rio.
 [X]  [X] N�o requer privil�gios de administrador.
#>

[CmdletBinding()]
param(
 [X]  [X] [Parameter()]
 [X]  [X] [string]$SourcePath = [OK]"",
 [X]  [X] 
 [X]  [X] [Parameter()]
 [X]  [X] [switch]$Force,
 [X]  [X] 
 [X]  [X] [Parameter()]
 [X]  [X] [switch]$NoBackup,
 [X]  [X] 
 [X]  [X] [Parameter()]
 [X]  [X] [switch]$SkipCustomizations,
 [X]  [X] 
 [X]  [X] [Parameter(DontShow)]
 [X]  [X] [switch]$BypassedExecution
)

# Define o caminho padr�o como a pasta onde o script est� localizado
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
 [X]  [X] $SourcePath = $PSScriptRoot
 [X]  [X] if ([string]::IsNullOrWhiteSpace($SourcePath)) {
 [X]  [X]  [X]  [X] # Fallback se PSScriptRoot n�o estiver dispon�vel
 [X]  [X]  [X]  [X] $SourcePath = Split-Path -Parent $MyInvocation.MyCommand.Path
 [X]  [X] }
}

# =============================================================================
# AUTO-RELAN�AMENTO COM BYPASS DE EXECU��O
# =============================================================================
# Este bloco garante que o script seja executado com a pol�tica de execu��o
# adequada, sem modificar permanentemente as configura��es do sistema.
# Extremamente seguro: apenas este script � executado com bypass tempor�rio.
# =============================================================================

if (-not $BypassedExecution) {
 [X]  [X] Write-Host [OK]"[LOCK] Verificando pol�tica de execu��o..." -ForegroundColor Cyan
 [X]  [X] 
 [X]  [X] # Captura a pol�tica atual para documenta��o no log
 [X]  [X] $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
 [X]  [X] Write-Host [OK]" [X]  Pol�tica atual (CurrentUser): $currentPolicy" -ForegroundColor Gray
 [X]  [X] 
 [X]  [X] # Verifica se precisa de bypass
 [X]  [X] $needsBypass = $false
 [X]  [X] try {
 [X]  [X]  [X]  [X] # Tenta uma opera��o de script simples
 [X]  [X]  [X]  [X] $null = [ScriptBlock]::Create("1 + 1").Invoke()
 [X]  [X] }
 [X]  [X] catch [System.Management.Automation.PSSecurityException] {
 [X]  [X]  [X]  [X] $needsBypass = $true
 [X]  [X] }
 [X]  [X] 
 [X]  [X] if ($needsBypass -or $currentPolicy -eq [OK]"Restricted" -or $currentPolicy -eq [OK]"AllSigned") {
 [X]  [X]  [X]  [X] Write-Host [OK]"[!] [X] Pol�tica de execu��o restritiva detectada." -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]"[SYNC] Relan�ando script com bypass tempor�rio..." -ForegroundColor Cyan
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"[i] [X] SEGURAN�A:" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Apenas ESTE script ser� executado com bypass" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � A pol�tica do sistema N�O ser� alterada" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � O bypass expira quando o script terminar" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Nenhum privil�gio de administrador � usado" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Constr�i argumentos para o relan�amento
 [X]  [X]  [X]  [X] $arguments = @(
 [X]  [X]  [X]  [X]  [X]  [X] [OK]"-ExecutionPolicy", [OK]"Bypass",
 [X]  [X]  [X]  [X]  [X]  [X] [OK]"-NoProfile",
 [X]  [X]  [X]  [X]  [X]  [X] [OK]"-File", [OK]"`"$PSCommandPath`"",
 [X]  [X]  [X]  [X]  [X]  [X] [OK]"-BypassedExecution"
 [X]  [X]  [X]  [X] )
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Adiciona par�metros originais
 [X]  [X]  [X]  [X] # SourcePath � sempre definido automaticamente, ent�o n�o precisa passar
 [X]  [X]  [X]  [X] if ($Force) {
 [X]  [X]  [X]  [X]  [X]  [X] $arguments += [OK]"-Force"
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] if ($NoBackup) {
 [X]  [X]  [X]  [X]  [X]  [X] $arguments += [OK]"-NoBackup"
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Relan�a o script com bypass tempor�rio
 [X]  [X]  [X]  [X] $processInfo = Start-Process -FilePath [OK]"powershell.exe" `
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  -ArgumentList $arguments `
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  -Wait `
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  -NoNewWindow `
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  -PassThru
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Retorna o c�digo de sa�da do processo relan�ado
 [X]  [X]  [X]  [X] exit $processInfo.ExitCode
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Host [OK]"[OK] Pol�tica de execu��o adequada: $currentPolicy" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X] }
}
else {
 [X]  [X] Write-Host [OK]"[OK] Executando com bypass tempor�rio (seguro)" -ForegroundColor Green
 [X]  [X] Write-Host [OK]""
}

# =============================================================================
# CONFIGURA��ES E CONSTANTES
# =============================================================================

$ErrorActionPreference = [OK]"Stop"
$script:LogFile = $null
$script:WarningCount = 0
$script:ErrorCount = 0
$script:SuccessCount = 0

# Cores para output
$ColorSuccess = [OK]"Green"
$ColorWarning = [OK]"Yellow"
$ColorError = [OK]"Red"
$ColorInfo = [OK]"Cyan"

# =============================================================================
# FUN��ES DE LOG
# =============================================================================

function Initialize-LogFile {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Inicializa o arquivo de log.
 [X]  [X] #>
 [X]  [X] try {
 [X]  [X]  [X]  [X] $logDir = Join-Path $env:USERPROFILE [OK]"CHAINSAW\logs"
 [X]  [X]  [X]  [X] if (-not (Test-Path $logDir)) {
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $logDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $timestamp = Get-Date -Format [OK]"yyyyMMdd_HHmmss"
 [X]  [X]  [X]  [X] $script:LogFile = Join-Path $logDir [OK]"install_$timestamp.log"
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $header = @"
================================================================================
CHAINSAW - Log de Instala��o
================================================================================
Data/Hora In�cio: $(Get-Date -Format [OK]"dd/MM/yyyy HH:mm:ss")
Usu�rio: $env:USERNAME
Computador: $env:COMPUTERNAME
Sistema: $([Environment]::OSVersion.VersionString)
PowerShell: $($PSVersionTable.PSVersion)
Caminho de Origem: $SourcePath
================================================================================

"@
 [X]  [X]  [X]  [X] Add-Content -Path $script:LogFile -Value $header
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Warning [OK]"N�o foi poss�vel criar arquivo de log: $_"
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Write-Log {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Escreve mensagem no log e na tela.
 [X]  [X] #>
 [X]  [X] param(
 [X]  [X]  [X]  [X] [Parameter(Mandatory)]
 [X]  [X]  [X]  [X] [string]$Message,
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] [Parameter()]
 [X]  [X]  [X]  [X] [ValidateSet("INFO", [OK]"SUCCESS", [OK]"WARNING", [OK]"ERROR")]
 [X]  [X]  [X]  [X] [string]$Level = [OK]"INFO",
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] [Parameter()]
 [X]  [X]  [X]  [X] [switch]$NoConsole
 [X]  [X] )
 [X]  [X] 
 [X]  [X] $timestamp = Get-Date -Format [OK]"yyyy-MM-dd HH:mm:ss"
 [X]  [X] $logEntry = [OK]"[$timestamp] [$Level] $Message"
 [X]  [X] 
 [X]  [X] # Escreve no arquivo de log
 [X]  [X] if ($script:LogFile) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] # Ignora erros de escrita no log para n�o interromper o processo
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Escreve no console
 [X]  [X] if (-not $NoConsole) {
 [X]  [X]  [X]  [X] switch ($Level) {
 [X]  [X]  [X]  [X]  [X]  [X] [OK]"SUCCESS" {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[OK] $Message" -ForegroundColor $ColorSuccess
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $script:SuccessCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] [OK]"WARNING" {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[!] $Message" -ForegroundColor $ColorWarning
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $script:WarningCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] [OK]"ERROR" {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[X] $Message" -ForegroundColor $ColorError
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $script:ErrorCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] default {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[i] $Message" -ForegroundColor $ColorInfo
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X] }
}

# =============================================================================
# FUN��ES DE VALIDA��O
# =============================================================================

function Test-Prerequisites {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Verifica pr�-requisitos para instala��o.
 [X]  [X] #>
 [X]  [X] Write-Log [OK]"Verificando pr�-requisitos..." -Level INFO
 [X]  [X] 
 [X]  [X] $allOk = $true
 [X]  [X] 
 [X]  [X] # Verifica vers�o do Windows
 [X]  [X] $osVersion = [Environment]::OSVersion.Version
 [X]  [X] if ($osVersion.Major -lt 10) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Windows 10 ou superior � necess�rio. Vers�o detectada: $($osVersion.ToString())" -Level ERROR
 [X]  [X]  [X]  [X] $allOk = $false
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Sistema operacional: Windows $($osVersion.Major).$($osVersion.Minor) ?" -Level SUCCESS
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Verifica vers�o do PowerShell
 [X]  [X] $psVersion = $PSVersionTable.PSVersion
 [X]  [X] if ($psVersion.Major -lt 5) {
 [X]  [X]  [X]  [X] Write-Log [OK]"PowerShell 5.1 ou superior � necess�rio. Vers�o detectada: $($psVersion.ToString())" -Level ERROR
 [X]  [X]  [X]  [X] $allOk = $false
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"PowerShell vers�o: $($psVersion.ToString()) ?" -Level SUCCESS
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Verifica acesso ao caminho de rede
 [X]  [X] Write-Log [OK]"Verificando acesso ao caminho de rede: $SourcePath" -Level INFO
 [X]  [X] if (-not (Test-Path $SourcePath)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"N�o foi poss�vel acessar o caminho de rede: $SourcePath" -Level ERROR
 [X]  [X]  [X]  [X] Write-Log [OK]"Verifique se voc� est� conectado � rede e tem permiss�es de acesso." -Level ERROR
 [X]  [X]  [X]  [X] $allOk = $false
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Acesso ao caminho de rede confirmado ?" -Level SUCCESS
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Verifica permiss�es de escrita no perfil do usu�rio
 [X]  [X] $testFile = Join-Path $env:USERPROFILE [OK]"CHAINSAW_test_$(Get-Date -Format 'yyyyMMddHHmmss').tmp"
 [X]  [X] try {
 [X]  [X]  [X]  [X] [System.IO.File]::WriteAllText($testFile, [OK]"test")
 [X]  [X]  [X]  [X] Remove-Item $testFile -Force -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] Write-Log [OK]"Permiss�es de escrita no perfil do usu�rio confirmadas ?" -Level SUCCESS
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Sem permiss�es de escrita no perfil do usu�rio: $env:USERPROFILE" -Level ERROR
 [X]  [X]  [X]  [X] $allOk = $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] return $allOk
}

function Test-SourceFiles {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Verifica se os arquivos de origem existem.
 [X]  [X] #>
 [X]  [X] param(
 [X]  [X]  [X]  [X] [ref]$SourceStampFile,
 [X]  [X]  [X]  [X] [ref]$SourceTemplatesFolder
 [X]  [X] )
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Verificando arquivos de origem..." -Level INFO
 [X]  [X] 
 [X]  [X] $allOk = $true
 [X]  [X] 
 [X]  [X] # Verifica arquivo stamp.png
 [X]  [X] $stampPath = Join-Path $SourcePath [OK]"assets\stamp.png"
 [X]  [X] if (Test-Path $stampPath) {
 [X]  [X]  [X]  [X] $SourceStampFile.Value = $stampPath
 [X]  [X]  [X]  [X] Write-Log [OK]"Arquivo stamp.png encontrado ?" -Level SUCCESS
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Arquivo n�o encontrado: $stampPath" -Level ERROR
 [X]  [X]  [X]  [X] $allOk = $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Verifica pasta Templates
 [X]  [X] $templatesPath = Join-Path $SourcePath [OK]"configs\Templates"
 [X]  [X] if (Test-Path $templatesPath) {
 [X]  [X]  [X]  [X] $SourceTemplatesFolder.Value = $templatesPath
 [X]  [X]  [X]  [X] Write-Log [OK]"Pasta Templates encontrada ?" -Level SUCCESS
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Pasta n�o encontrada: $templatesPath" -Level ERROR
 [X]  [X]  [X]  [X] $allOk = $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] return $allOk
}

# =============================================================================
# FUN��ES AUXILIARES
# =============================================================================

function Test-WordRunning {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Verifica se o Microsoft Word est� em execu��o.
 [X]  [X] #>
 [X]  [X] $wordProcesses = Get-Process -Name [OK]"WINWORD" -ErrorAction SilentlyContinue
 [X]  [X] return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

# =============================================================================
# FUN��ES DE BACKUP
# =============================================================================

function Backup-TemplatesFolder {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Cria backup da pasta Templates existente.
 [X]  [X] #>
 [X]  [X] param(
 [X]  [X]  [X]  [X] [Parameter(Mandatory)]
 [X]  [X]  [X]  [X] [string]$SourceFolder
 [X]  [X] )
 [X]  [X] 
 [X]  [X] if (-not (Test-Path $SourceFolder)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Pasta Templates n�o existe, backup n�o necess�rio." -Level INFO
 [X]  [X]  [X]  [X] return $null
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Verifica se o Word est� aberto
 [X]  [X] if (Test-WordRunning) {
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]"� [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] [!] MICROSOFT WORD ABERTO [!] [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] �" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"O Microsoft Word est� em execu��o e deve ser fechado antes de" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]"continuar com a instala��o." -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"Por favor:" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] 1. Salve todos os documentos abertos no Word" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] 2. Feche completamente o Microsoft Word" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] 3. Pressione qualquer tecla para continuar" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Aguardando fechamento do Word..." -Level WARNING
 [X]  [X]  [X]  [X] $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Verifica novamente
 [X]  [X]  [X]  [X] if (Test-WordRunning) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Word ainda est� aberto - abortando instala��o" -Level ERROR
 [X]  [X]  [X]  [X]  [X]  [X] throw [OK]"Microsoft Word deve ser fechado antes da instala��o."
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Host [OK]"? Word fechado, continuando..." -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X] }
 [X]  [X] 
 [X]  [X] $timestamp = Get-Date -Format [OK]"yyyyMMdd_HHmmss"
 [X]  [X] $backupName = [OK]"Templates_backup_$timestamp"
 [X]  [X] $backupPath = Join-Path (Split-Path $SourceFolder -Parent) $backupName
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Criando backup da pasta Templates..." -Level INFO
 [X]  [X] Write-Log [OK]"Origem: $SourceFolder" -Level INFO
 [X]  [X] Write-Log [OK]"Destino: $backupPath" -Level INFO
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] # Tenta usar Rename-Item primeiro (mais r�pido)
 [X]  [X]  [X]  [X] Rename-Item -Path $SourceFolder -NewName $backupName -Force -ErrorAction Stop
 [X]  [X]  [X]  [X] Write-Log [OK]"Backup criado com sucesso: $backupName ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $backupPath
 [X]  [X] }
 [X]  [X] catch [System.IO.IOException] {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro de acesso ao renomear (poss�vel arquivo em uso)" -Level WARNING
 [X]  [X]  [X]  [X] Write-Log [OK]"Tentando m�todo alternativo (c�pia)..." -Level INFO
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] # M�todo alternativo: copiar e depois deletar
 [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $SourceFolder -Destination $backupPath -Recurse -Force -ErrorAction Stop
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] # Aguarda um pouco para liberar arquivos
 [X]  [X]  [X]  [X]  [X]  [X] Start-Sleep -Seconds 1
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] # Remove a pasta original
 [X]  [X]  [X]  [X]  [X]  [X] Remove-Item -Path $SourceFolder -Recurse -Force -ErrorAction Stop
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Backup criado com sucesso (m�todo c�pia): $backupName ?" -Level SUCCESS
 [X]  [X]  [X]  [X]  [X]  [X] return $backupPath
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao criar backup com m�todo alternativo: $_" -Level ERROR
 [X]  [X]  [X]  [X]  [X]  [X] throw [OK]"N�o foi poss�vel criar backup. Certifique-se de que o Word est� fechado e que n�o h� arquivos em uso na pasta Templates."
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao criar backup: $_" -Level ERROR
 [X]  [X]  [X]  [X] throw
 [X]  [X] }
}

function Remove-OldBackups {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Remove backups antigos mantendo apenas os mais recentes.
 [X]  [X] #>
 [X]  [X] param(
 [X]  [X]  [X]  [X] [Parameter(Mandatory)]
 [X]  [X]  [X]  [X] [string]$BackupFolder,
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] [Parameter()]
 [X]  [X]  [X]  [X] [int]$KeepCount = 5
 [X]  [X] )
 [X]  [X] 
 [X]  [X] $backupParent = Split-Path $BackupFolder -Parent
 [X]  [X] $backups = Get-ChildItem -Path $backupParent -Directory -Filter [OK]"Templates_backup_*" |
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  Sort-Object Name -Descending
 [X]  [X] 
 [X]  [X] if ($backups.Count -gt $KeepCount) {
 [X]  [X]  [X]  [X] $toRemove = $backups | Select-Object -Skip $KeepCount
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Removendo backups antigos (mantendo os $KeepCount mais recentes)..." -Level INFO
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] foreach ($backup in $toRemove) {
 [X]  [X]  [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Remove-Item -Path $backup.FullName -Recurse -Force -ErrorAction Stop
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Backup removido: $($backup.Name)" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao remover backup $($backup.Name): $_" -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X] }
}

# =============================================================================
# FUN��ES DE INSTALA��O
# =============================================================================

function Copy-StampFile {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Copia o arquivo stamp.png para a pasta do usu�rio.
 [X]  [X] #>
 [X]  [X] param(
 [X]  [X]  [X]  [X] [Parameter(Mandatory)]
 [X]  [X]  [X]  [X] [string]$SourceFile
 [X]  [X] )
 [X]  [X] 
 [X]  [X] $destFolder = Join-Path $env:USERPROFILE [OK]"CHAINSAW\assets"
 [X]  [X] $destFile = Join-Path $destFolder [OK]"stamp.png"
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Copiando arquivo stamp.png..." -Level INFO
 [X]  [X] Write-Log [OK]"Origem: $SourceFile" -Level INFO
 [X]  [X] Write-Log [OK]"Destino: $destFile" -Level INFO
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] # Verifica se origem e destino s�o o mesmo arquivo
 [X]  [X]  [X]  [X] $sourceFullPath = (Resolve-Path $SourceFile).Path
 [X]  [X]  [X]  [X] $destFullPath = if (Test-Path $destFile) { (Resolve-Path $destFile).Path } else { $null }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if ($sourceFullPath -eq $destFullPath) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Arquivo j� est� no local correto (origem = destino), pulando c�pia" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Arquivo stamp.png j� est� instalado ?" -Level SUCCESS
 [X]  [X]  [X]  [X]  [X]  [X] return $true
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Cria pasta de destino se n�o existir
 [X]  [X]  [X]  [X] if (-not (Test-Path $destFolder)) {
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Pasta criada: $destFolder" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Copia o arquivo
 [X]  [X]  [X]  [X] Copy-Item -Path $SourceFile -Destination $destFile -Force -ErrorAction Stop
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Verifica se o arquivo foi copiado corretamente
 [X]  [X]  [X]  [X] if (Test-Path $destFile) {
 [X]  [X]  [X]  [X]  [X]  [X] $sourceSize = (Get-Item $SourceFile).Length
 [X]  [X]  [X]  [X]  [X]  [X] $destSize = (Get-Item $destFile).Length
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] if ($sourceSize -eq $destSize) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Arquivo stamp.png copiado com sucesso ?" -Level SUCCESS
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] return $true
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Tamanhos diferentes: origem=$sourceSize, destino=$destSize" -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] return $false
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Arquivo n�o foi copiado corretamente" -Level ERROR
 [X]  [X]  [X]  [X]  [X]  [X] return $false
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao copiar stamp.png: $_" -Level ERROR
 [X]  [X]  [X]  [X] throw
 [X]  [X] }
}

function Copy-TemplatesFolder {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Copia a pasta Templates da rede para o perfil do usu�rio.
 [X]  [X] #>
 [X]  [X] param(
 [X]  [X]  [X]  [X] [Parameter(Mandatory)]
 [X]  [X]  [X]  [X] [string]$SourceFolder,
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] [Parameter(Mandatory)]
 [X]  [X]  [X]  [X] [string]$DestFolder
 [X]  [X] )
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Copiando pasta Templates..." -Level INFO
 [X]  [X] Write-Log [OK]"Origem: $SourceFolder" -Level INFO
 [X]  [X] Write-Log [OK]"Destino: $DestFolder" -Level INFO
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] # Verifica se origem e destino s�o o mesmo local
 [X]  [X]  [X]  [X] $sourceFullPath = (Resolve-Path $SourceFolder).Path.TrimEnd('\')
 [X]  [X]  [X]  [X] $destFullPath = if (Test-Path $DestFolder) { (Resolve-Path $DestFolder).Path.TrimEnd('\') } else { $null }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if ($sourceFullPath -eq $destFullPath) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"A pasta Templates j� est� no local correto (origem = destino), pulando c�pia" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Pasta Templates j� est� instalada ?" -Level SUCCESS
 [X]  [X]  [X]  [X]  [X]  [X] return $true
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Cria pasta de destino
 [X]  [X]  [X]  [X] if (-not (Test-Path $DestFolder)) {
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $DestFolder -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Copia todos os arquivos e subpastas
 [X]  [X]  [X]  [X] $itemsToCopy = Get-ChildItem -Path $SourceFolder -Recurse
 [X]  [X]  [X]  [X] $totalItems = $itemsToCopy.Count
 [X]  [X]  [X]  [X] $copiedItems = 0
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Total de itens a copiar: $totalItems" -Level INFO
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] foreach ($item in $itemsToCopy) {
 [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $item.FullName.Substring($SourceFolder.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X] $destPath = Join-Path $DestFolder $relativePath
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] if ($item.PSIsContainer) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] # � uma pasta
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destPath)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] # � um arquivo
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destPath -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $item.FullName -Destination $destPath -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $copiedItems++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] # Progress
 [X]  [X]  [X]  [X]  [X]  [X] if ($copiedItems % 10 -eq 0) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Progress -Activity [OK]"Copiando Templates" -Status [OK]"$copiedItems de $totalItems arquivos copiados" -PercentComplete (($copiedItems / $totalItems) * 100)
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Progress -Activity [OK]"Copiando Templates" -Completed
 [X]  [X]  [X]  [X] Write-Log [OK]"Pasta Templates copiada com sucesso ($copiedItems arquivos) ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao copiar pasta Templates: $_" -Level ERROR
 [X]  [X]  [X]  [X] throw
 [X]  [X] }
}

# =============================================================================
# FUN��ES DE IMPORTA��O DE PERSONALIZA��ES
# =============================================================================

function Test-CustomizationsAvailable {
 [X]  [X] param([string]$ImportPath)
 [X]  [X] 
 [X]  [X] if (-not (Test-Path $ImportPath)) {
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Verifica se h� um manifesto ou arquivos para importar
 [X]  [X] $manifestPath = Join-Path $ImportPath [OK]"MANIFEST.json"
 [X]  [X] $hasManifest = Test-Path $manifestPath
 [X]  [X] 
 [X]  [X] if ($hasManifest) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Manifesto encontrado: $($manifest.TotalItems) itens" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Exportado em: $($manifest.ExportDate) por $($manifest.UserName)" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao ler manifesto: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] return $true
}

function Backup-WordCustomizations {
 [X]  [X] param([string]$BackupReason = [OK]"pr�-importa��o")
 [X]  [X] 
 [X]  [X] if ($NoBackup) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Backup de personaliza��es desabilitado (-NoBackup)" -Level WARNING
 [X]  [X]  [X]  [X] return $null
 [X]  [X] }
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Criando backup das personaliza��es do Word ($BackupReason)..." -Level INFO
 [X]  [X] 
 [X]  [X] $timestamp = Get-Date -Format [OK]"yyyyMMdd_HHmmss"
 [X]  [X] $backupPath = Join-Path $env:USERPROFILE [OK]"CHAINSAW\backups\word-customizations_$timestamp"
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] if (-not (Test-Path $backupPath)) {
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $templatesPath = Join-Path $env:APPDATA [OK]"Microsoft\Templates"
 [X]  [X]  [X]  [X] $localAppDataPath = $env:LOCALAPPDATA
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Backup do Normal.dotm
 [X]  [X]  [X]  [X] $normalPath = Join-Path $templatesPath [OK]"Normal.dotm"
 [X]  [X]  [X]  [X] if (Test-Path $normalPath) {
 [X]  [X]  [X]  [X]  [X]  [X] $destNormal = Join-Path $backupPath [OK]"Templates"
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destNormal -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $normalPath -Destination $destNormal -Force
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Normal.dotm backup criado" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Backup de personaliza��es UI
 [X]  [X]  [X]  [X] $uiPath = Join-Path $localAppDataPath [OK]"Microsoft\Office"
 [X]  [X]  [X]  [X] $uiFiles = Get-ChildItem -Path $uiPath -Filter [OK]"*.officeUI" -Recurse -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] if ($uiFiles.Count -gt 0) {
 [X]  [X]  [X]  [X]  [X]  [X] $destUI = Join-Path $backupPath [OK]"OfficeCustomUI"
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destUI -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $uiFiles) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination (Join-Path $destUI $file.Name) -Force
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Personaliza��es UI backup criado: $($uiFiles.Count) arquivos" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Backup de personaliza��es criado em: $backupPath ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $backupPath
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao criar backup de personaliza��es: $_" -Level ERROR
 [X]  [X]  [X]  [X] return $null
 [X]  [X] }
}

function Import-NormalTemplate {
 [X]  [X] param([string]$ImportPath)
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Importando Normal.dotm..." -Level INFO
 [X]  [X] 
 [X]  [X] $sourcePath = Join-Path $ImportPath [OK]"Templates\Normal.dotm"
 [X]  [X] $templatesPath = Join-Path $env:APPDATA [OK]"Microsoft\Templates"
 [X]  [X] $destPath = Join-Path $templatesPath [OK]"Normal.dotm"
 [X]  [X] 
 [X]  [X] if (-not (Test-Path $sourcePath)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Normal.dotm n�o encontrado no pacote de importa��o" -Level WARNING
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] if (-not (Test-Path $templatesPath)) {
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $templatesPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Copy-Item -Path $sourcePath -Destination $destPath -Force
 [X]  [X]  [X]  [X] Write-Log [OK]"Normal.dotm importado com sucesso ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao importar Normal.dotm: $_" -Level ERROR
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Import-BuildingBlocks {
 [X]  [X] param([string]$ImportPath)
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Importando Building Blocks..." -Level INFO
 [X]  [X] 
 [X]  [X] $templatesPath = Join-Path $env:APPDATA [OK]"Microsoft\Templates"
 [X]  [X] $sourceManaged = Join-Path $ImportPath [OK]"Templates\LiveContent\16\Managed\Word Document Building Blocks"
 [X]  [X] $sourceUser = Join-Path $ImportPath [OK]"Templates\LiveContent\16\User\Word Document Building Blocks"
 [X]  [X] 
 [X]  [X] $destManaged = Join-Path $templatesPath [OK]"LiveContent\16\Managed\Word Document Building Blocks"
 [X]  [X] $destUser = Join-Path $templatesPath [OK]"LiveContent\16\User\Word Document Building Blocks"
 [X]  [X] 
 [X]  [X] $importedCount = 0
 [X]  [X] 
 [X]  [X] # Importa Building Blocks gerenciados
 [X]  [X] if (Test-Path $sourceManaged) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destManaged)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destManaged -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $sourceManaged -Recurse -File
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $file.FullName.Substring($sourceManaged.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destManaged $relativePath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destFile -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $importedCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Building Blocks gerenciados importados: $($files.Count) arquivos" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao importar Building Blocks gerenciados: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Importa Building Blocks do usu�rio
 [X]  [X] if (Test-Path $sourceUser) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destUser)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destUser -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $sourceUser -Recurse -File
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $file.FullName.Substring($sourceUser.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destUser $relativePath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destFile -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $importedCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Building Blocks do usu�rio importados: $($files.Count) arquivos" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao importar Building Blocks do usu�rio: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] if ($importedCount -gt 0) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Building Blocks importados: $importedCount arquivos ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum Building Block para importar" -Level INFO
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Import-DocumentThemes {
 [X]  [X] param([string]$ImportPath)
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Importando temas de documentos..." -Level INFO
 [X]  [X] 
 [X]  [X] $templatesPath = Join-Path $env:APPDATA [OK]"Microsoft\Templates"
 [X]  [X] $sourceManaged = Join-Path $ImportPath [OK]"Templates\LiveContent\16\Managed\Document Themes"
 [X]  [X] $sourceUser = Join-Path $ImportPath [OK]"Templates\LiveContent\16\User\Document Themes"
 [X]  [X] 
 [X]  [X] $destManaged = Join-Path $templatesPath [OK]"LiveContent\16\Managed\Document Themes"
 [X]  [X] $destUser = Join-Path $templatesPath [OK]"LiveContent\16\User\Document Themes"
 [X]  [X] 
 [X]  [X] $importedCount = 0
 [X]  [X] 
 [X]  [X] # Temas gerenciados
 [X]  [X] if (Test-Path $sourceManaged) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destManaged)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destManaged -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $sourceManaged -Recurse -File
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $file.FullName.Substring($sourceManaged.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destManaged $relativePath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destFile -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $importedCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao importar temas gerenciados: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Temas do usu�rio
 [X]  [X] if (Test-Path $sourceUser) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destUser)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destUser -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $sourceUser -Recurse -File
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $file.FullName.Substring($sourceUser.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destUser $relativePath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destFile -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $importedCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao importar temas do usu�rio: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] if ($importedCount -gt 0) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Temas importados: $importedCount arquivos ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum tema para importar" -Level INFO
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Import-RibbonCustomization {
 [X]  [X] param([string]$ImportPath)
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Importando personaliza��o da Faixa de Op��es..." -Level INFO
 [X]  [X] 
 [X]  [X] $sourcePath = Join-Path $ImportPath [OK]"RibbonCustomization"
 [X]  [X] 
 [X]  [X] if (-not (Test-Path $sourcePath)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhuma personaliza��o do Ribbon para importar" -Level INFO
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $sourcePath -Filter [OK]"*.officeUI" -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if ($files.Count -eq 0) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum arquivo de personaliza��o Ribbon encontrado" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] return $false
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X] # Tenta os locais poss�veis
 [X]  [X]  [X]  [X]  [X]  [X] $possibleDests = @(
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] (Join-Path $env:LOCALAPPDATA [OK]"Microsoft\Office"),
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] (Join-Path $env:APPDATA [OK]"Microsoft\Office")
 [X]  [X]  [X]  [X]  [X]  [X] )
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($destPath in $possibleDests) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destPath)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destPath $file.Name
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Ribbon importado para: $destFile" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Personaliza��o do Ribbon importada: $($files.Count) arquivos ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao importar Ribbon: $_" -Level ERROR
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Import-OfficeCustomUI {
 [X]  [X] param([string]$ImportPath)
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Importando personaliza��es da interface..." -Level INFO
 [X]  [X] 
 [X]  [X] $sourcePath = Join-Path $ImportPath [OK]"OfficeCustomUI"
 [X]  [X] 
 [X]  [X] if (-not (Test-Path $sourcePath)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhuma personaliza��o UI para importar" -Level INFO
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $sourcePath -Filter [OK]"*.officeUI" -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if ($files.Count -eq 0) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum arquivo de personaliza��o UI encontrado" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] return $false
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $destPath = Join-Path $env:LOCALAPPDATA [OK]"Microsoft\Office"
 [X]  [X]  [X]  [X] if (-not (Test-Path $destPath)) {
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destPath $file.Name
 [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Personaliza��es UI importadas: $($files.Count) arquivos ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao importar personaliza��es UI: $_" -Level ERROR
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Import-WordCustomizations {
 [X]  [X] param([string]$ImportPath)
 [X]  [X] 
 [X]  [X] Write-Log [OK]"=== Iniciando importa��o de personaliza��es ===" -Level INFO
 [X]  [X] 
 [X]  [X] # Verifica se o Word est� em execu��o
 [X]  [X] if (Test-WordRunning) {
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]"� [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] [!] MICROSOFT WORD ABERTO [!] [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] �" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"O Microsoft Word est� em execu��o e deve ser fechado antes de" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]"importar as personaliza��es." -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"Por favor:" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] 1. Salve todos os documentos abertos no Word" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] 2. Feche completamente o Microsoft Word" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] 3. Execute este script novamente" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Importa��o abortada: Word est� em execu��o" -Level WARNING
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Cria backup
 [X]  [X] $backupPath = Backup-WordCustomizations -BackupReason [OK]"pr�-importa��o de personaliza��es"
 [X]  [X] if ($null -eq $backupPath -and -not $NoBackup) {
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"[!] Falha ao criar backup das personaliza��es atuais." -ForegroundColor Yellow
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if (-not $Force) {
 [X]  [X]  [X]  [X]  [X]  [X] $response = Read-Host [OK]"Continuar mesmo assim? (S/N)"
 [X]  [X]  [X]  [X]  [X]  [X] if ($response -notmatch '^[Ss]$') {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Importa��o cancelada: falha no backup" -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] return $false
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X] Write-Host [OK]" [X] ETAPA 6: Importa��o de Personaliza��es do Word" -ForegroundColor White
 [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X] Write-Host [OK]""
 [X]  [X] 
 [X]  [X] # Importa��es
 [X]  [X] $importedCount = 0
 [X]  [X] 
 [X]  [X] if (Import-NormalTemplate -ImportPath $ImportPath) { $importedCount++ }
 [X]  [X] if (Import-BuildingBlocks -ImportPath $ImportPath) { $importedCount++ }
 [X]  [X] if (Import-DocumentThemes -ImportPath $ImportPath) { $importedCount++ }
 [X]  [X] if (Import-RibbonCustomization -ImportPath $ImportPath) { $importedCount++ }
 [X]  [X] if (Import-OfficeCustomUI -ImportPath $ImportPath) { $importedCount++ }
 [X]  [X] 
 [X]  [X] if ($importedCount -gt 0) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Total de categorias de personaliza��es importadas: $importedCount ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhuma personaliza��o foi importada" -Level WARNING
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

# =============================================================================
# GERENCIAMENTO DO WORD
# =============================================================================

function Test-WordRunning {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Verifica se h� processos do Word em execu��o.
 [X]  [X] #>
 [X]  [X] $wordProcesses = Get-Process -Name [OK]"WINWORD" -ErrorAction SilentlyContinue
 [X]  [X] return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Stop-WordProcesses {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Fecha for�adamente todos os processos do Word.
 [X]  [X] .DESCRIPTION
 [X]  [X]  [X]  [X] Encerra apenas processos WINWORD.EXE (Microsoft Word), sem afetar
 [X]  [X]  [X]  [X] outros aplicativos do Office como Excel (EXCEL.EXE) ou PowerPoint (POWERPNT.EXE).
 [X]  [X] #>
 [X]  [X] param(
 [X]  [X]  [X]  [X] [Parameter()]
 [X]  [X]  [X]  [X] [switch]$Force
 [X]  [X] )
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] $wordProcesses = Get-Process -Name [OK]"WINWORD" -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if ($null -eq $wordProcesses -or $wordProcesses.Count -eq 0) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum processo do Word em execu��o" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] return $true
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Encontrados $($wordProcesses.Count) processo(s) do Word em execu��o" -Level INFO
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] foreach ($process in $wordProcesses) {
 [X]  [X]  [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Encerrando processo Word (PID: $($process.Id))..." -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if ($Force) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] # Encerra for�adamente
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $process.Kill()
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $process.WaitForExit(5000) # Aguarda at� 5 segundos
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] # Tenta encerrar graciosamente primeiro
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $process.CloseMainWindow() | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Start-Sleep -Milliseconds 500
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not $process.HasExited) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $process.Kill()
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $process.WaitForExit(5000)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Processo Word (PID: $($process.Id)) encerrado com sucesso" -Level SUCCESS
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao encerrar processo Word (PID: $($process.Id)): $_" -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Aguarda um momento e verifica se todos foram fechados
 [X]  [X]  [X]  [X] Start-Sleep -Milliseconds 1000
 [X]  [X]  [X]  [X] $remainingProcesses = Get-Process -Name [OK]"WINWORD" -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if ($null -ne $remainingProcesses -and $remainingProcesses.Count -gt 0) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Ainda h� $($remainingProcesses.Count) processo(s) do Word em execu��o" -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X] return $false
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Todos os processos do Word foram encerrados com sucesso" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao encerrar processos do Word: $_" -Level ERROR
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Confirm-CloseWord {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Solicita que o usu�rio salve e feche o Word, ou cancela a opera��o.
 [X]  [X] .DESCRIPTION
 [X]  [X]  [X]  [X] Exibe aviso ao usu�rio e aguarda confirma��o antes de fechar o Word for�adamente.
 [X]  [X]  [X]  [X] Retorna $true se o usu�rio autorizar, $false se cancelar.
 [X]  [X] #>
 [X]  [X] 
 [X]  [X] # Verifica se Word est� em execu��o
 [X]  [X] if (-not (Test-WordRunning)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Word n�o est� em execu��o - prosseguindo..." -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] 
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]"� [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] [!] ATEN��O [!] [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] �" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"O Microsoft Word est� atualmente em execu��o!" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"IMPORTANTE:" -ForegroundColor Red
 [X]  [X] Write-Host [OK]" [X] � SALVE todos os seus documentos abertos no Word" -ForegroundColor White
 [X]  [X] Write-Host [OK]" [X] � FECHE o Word completamente" -ForegroundColor White
 [X]  [X] Write-Host [OK]" [X] � Outros aplicativos do Office (Excel, PowerPoint) N�O ser�o afetados" -ForegroundColor Gray
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"Se voc� continuar, o Word ser� FECHADO FOR�ADAMENTE e" -ForegroundColor Red
 [X]  [X] Write-Host [OK]"qualquer trabalho n�o salvo SER� PERDIDO!" -ForegroundColor Red
 [X]  [X] Write-Host [OK]""
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Word em execu��o - solicitando confirma��o do usu�rio" -Level WARNING
 [X]  [X] 
 [X]  [X] # Aguarda confirma��o
 [X]  [X] $response = Read-Host [OK]"Deseja FECHAR o Word e continuar a instala��o? (S/N)"
 [X]  [X] 
 [X]  [X] if ($response -notmatch '^[Ss]$') {
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"? Instala��o cancelada pelo usu�rio" -ForegroundColor Cyan
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] Salve seus documentos e execute o script novamente quando estiver pronto." -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Log [OK]"Instala��o cancelada - usu�rio optou por n�o fechar o Word" -Level WARNING
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Usu�rio confirmou - fecha o Word
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"Fechando Microsoft Word..." -ForegroundColor Cyan
 [X]  [X] Write-Log [OK]"Usu�rio autorizou o fechamento do Word" -Level INFO
 [X]  [X] 
 [X]  [X] if (Stop-WordProcesses -Force) {
 [X]  [X]  [X]  [X] Write-Host [OK]"? Word fechado com sucesso" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] # Aguarda um pouco para garantir que recursos foram liberados
 [X]  [X]  [X]  [X] Start-Sleep -Seconds 2
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Host [OK]"? N�o foi poss�vel fechar o Word completamente" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Log [OK]"Falha ao fechar Word - cancelando instala��o" -Level ERROR
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $retry = Read-Host [OK]"Deseja tentar novamente? (S/N)"
 [X]  [X]  [X]  [X] if ($retry -match '^[Ss]$') {
 [X]  [X]  [X]  [X]  [X]  [X] return Confirm-CloseWord # Recurs�o para tentar novamente
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

# =============================================================================
# FUN��O PRINCIPAL
# =============================================================================

function Install-CHAINSAWConfig {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Fun��o principal de instala��o.
 [X]  [X] #>
 [X]  [X] 
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Cyan
 [X]  [X] Write-Host [OK]"� [X]  [X]  [X]  [X]  [X] CHAINSAW - Instala��o de Configura��es do Word [X]  [X]  [X]  �" -ForegroundColor Cyan
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Cyan
 [X]  [X] Write-Host [OK]""
 [X]  [X] 
 [X]  [X] # Inicializa log
 [X]  [X] if (-not (Initialize-LogFile)) {
 [X]  [X]  [X]  [X] Write-Warning [OK]"Continuando sem arquivo de log..."
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Host [OK]"[LOG] Arquivo de log: $script:LogFile" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X] }
 [X]  [X] 
 [X]  [X] $startTime = Get-Date
 [X]  [X] Write-Log [OK]"=== IN�CIO DA INSTALA��O ===" -Level INFO
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] # 0. Verificar e fechar Word se necess�rio
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] ETAPA 0: Verifica��o do Microsoft Word" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if (-not (Confirm-CloseWord)) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Instala��o cancelada - Word n�o foi fechado" -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X] return
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 1. Verificar pr�-requisitos
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] ETAPA 1: Verifica��o de Pr�-requisitos" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if (-not (Test-Prerequisites)) {
 [X]  [X]  [X]  [X]  [X]  [X] throw [OK]"Pr�-requisitos n�o atendidos. Verifique os erros acima."
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 2. Verificar arquivos de origem
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] ETAPA 2: Verifica��o de Arquivos de Origem" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $sourceStampFile = $null
 [X]  [X]  [X]  [X] $sourceTemplatesFolder = $null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if (-not (Test-SourceFiles -SourceStampFile ([ref]$sourceStampFile) -SourceTemplatesFolder ([ref]$sourceTemplatesFolder))) {
 [X]  [X]  [X]  [X]  [X]  [X] throw [OK]"Arquivos de origem n�o encontrados. Verifique os erros acima."
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 3. Confirma��o do usu�rio
 [X]  [X]  [X]  [X] if (-not $Force) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] CONFIRMA��O" -ForegroundColor White
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"As seguintes opera��es ser�o realizadas:" -ForegroundColor Yellow
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] 1. Copiar stamp.png para: $env:USERPROFILE\CHAINSAW\assets\" -ForegroundColor White
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] 2. Fazer backup da pasta Templates atual (se existir)" -ForegroundColor White
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] 3. Copiar nova pasta Templates da rede" -ForegroundColor White
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $response = Read-Host [OK]"Deseja continuar? (S/N)"
 [X]  [X]  [X]  [X]  [X]  [X] if ($response -notmatch '^[Ss]$') {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Instala��o cancelada pelo usu�rio." -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] return
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 4. Copiar arquivo stamp.png
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] ETAPA 3: C�pia do Arquivo stamp.png" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Copy-StampFile -SourceFile $sourceStampFile | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 5. Backup da pasta Templates
 [X]  [X]  [X]  [X] $templatesPath = Join-Path $env:APPDATA [OK]"Microsoft\Templates"
 [X]  [X]  [X]  [X] $backupPath = $null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if (-not $NoBackup) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] ETAPA 4: Backup da Pasta Templates" -ForegroundColor White
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $backupPath = Backup-TemplatesFolder -SourceFolder $templatesPath
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] if ($backupPath) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Remove-OldBackups -BackupFolder $backupPath -KeepCount 5
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Backup desabilitado pelo par�metro -NoBackup" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 6. Copiar pasta Templates
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] ETAPA 5: C�pia da Pasta Templates" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Copy-TemplatesFolder -SourceFolder $sourceTemplatesFolder -DestFolder $templatesPath | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 7. Detectar e importar personaliza��es (se dispon�veis)
 [X]  [X]  [X]  [X] if (-not $SkipCustomizations) {
 [X]  [X]  [X]  [X]  [X]  [X] $exportedConfigPath = Join-Path $SourcePath [OK]"exported-config"
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] if (Test-CustomizationsAvailable -ImportPath $exportedConfigPath) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor Cyan
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] PERSONALIZA��ES DO WORD DETECTADAS!" -ForegroundColor White
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor Cyan
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[NEW] Personaliza��es exportadas foram encontradas em:" -ForegroundColor Cyan
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  $exportedConfigPath" -ForegroundColor Gray
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[PKG] Conte�do que ser� importado:" -ForegroundColor White
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Faixa de Op��es Personalizada (Ribbon)" -ForegroundColor Gray
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Partes R�pidas (Quick Parts)" -ForegroundColor Gray
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Blocos de Constru��o (Building Blocks)" -ForegroundColor Gray
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Temas de Documentos" -ForegroundColor Gray
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Template Normal.dotm" -ForegroundColor Gray
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $importCustomizations = $true
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not $Force) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $response = Read-Host [OK]"Deseja importar estas personaliza��es agora? (S/N)"
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $importCustomizations = ($response -match '^[Ss]$')
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if ($importCustomizations) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Iniciando importa��o de personaliza��es..." -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $imported = Import-WordCustomizations -ImportPath $exportedConfigPath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if ($imported) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[OK] Personaliza��es importadas com sucesso!" -ForegroundColor Green
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[i] IMPORTANTE:" -ForegroundColor Cyan
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  As personaliza��es ser�o vis�veis na pr�xima vez" -ForegroundColor Yellow
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  que voc� abrir o Microsoft Word." -ForegroundColor Yellow
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[!] Personaliza��es n�o foram importadas completamente." -ForegroundColor Yellow
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] Verifique o log para mais detalhes." -ForegroundColor Yellow
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[i] Importa��o de personaliza��es ignorada." -ForegroundColor Cyan
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] Para importar mais tarde, execute: .\install.ps1" -ForegroundColor Gray
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Importa��o de personaliza��es ignorada pelo usu�rio" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Pasta 'exported-config' n�o encontrada - pulando importa��o" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Importa��o de personaliza��es desabilitada (-SkipCustomizations)" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Sucesso!
 [X]  [X]  [X]  [X] $endTime = Get-Date
 [X]  [X]  [X]  [X] $duration = $endTime - $startTime
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]"� [X]  [X]  [X]  [X]  [X]  [X]  [X] INSTALA��O CONCLU�DA COM SUCESSO! [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  �" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"[CHART] Resumo da Instala��o:" -ForegroundColor Cyan
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Opera��es bem-sucedidas: $script:SuccessCount" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Avisos: $script:WarningCount" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Erros: $script:ErrorCount" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  � Tempo decorrido: $($duration.ToString('mm\:ss'))" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if ($backupPath) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[SAVE] Backup criado em:" -ForegroundColor Cyan
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X]  $backupPath" -ForegroundColor Gray
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Host [OK]"[LOG] Log completo salvo em:" -ForegroundColor Cyan
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  $script:LogFile" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"=== INSTALA��O CONCLU�DA COM SUCESSO ===" -Level SUCCESS
 [X]  [X]  [X]  [X] Write-Log [OK]"Dura��o: $($duration.ToString('mm\:ss'))" -Level INFO
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] $endTime = Get-Date
 [X]  [X]  [X]  [X] $duration = $endTime - $startTime
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]"� [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] ERRO NA INSTALA��O! [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  �" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"[X] Erro: $($_.Exception.Message)" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"[LOG] Verifique o arquivo de log para mais detalhes:" -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  $script:LogFile" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"=== INSTALA��O FALHOU ===" -Level ERROR
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro: $($_.Exception.Message)" -Level ERROR
 [X]  [X]  [X]  [X] Write-Log [OK]"Stack trace: $($_.ScriptStackTrace)" -Level ERROR
 [X]  [X]  [X]  [X] Write-Log [OK]"Dura��o at� falha: $($duration.ToString('mm\:ss'))" -Level INFO
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # Tenta reverter mudan�as se poss�vel
 [X]  [X]  [X]  [X] if ($backupPath -and (Test-Path $backupPath)) {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[SYNC] Tentando reverter mudan�as..." -ForegroundColor Yellow
 [X]  [X]  [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $templatesPath = Join-Path $env:APPDATA [OK]"Microsoft\Templates"
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (Test-Path $templatesPath) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Remove-Item -Path $templatesPath -Recurse -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Rename-Item -Path $backupPath -NewName [OK]"Templates" -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"[OK] Backup restaurado com sucesso" -ForegroundColor Green
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Backup restaurado ap�s falha na instala��o" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]"? N�o foi poss�vel restaurar o backup automaticamente" -ForegroundColor Red
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Host [OK]" [X] Backup dispon�vel em: $backupPath" -ForegroundColor Yellow
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Falha ao restaurar backup: $_" -Level ERROR
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] throw
 [X]  [X] }
}

# =============================================================================
# EXECU��O
# =============================================================================

# Verifica se o script est� sendo executado como administrador
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Red
 [X]  [X] Write-Host [OK]"� [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] [!] AVISO IMPORTANTE [!] [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] �" -ForegroundColor Red
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Red
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"[X] Este script est� sendo executado com privil�gios de Administrador." -ForegroundColor Red
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"[!] [X] PROBLEMA:" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]" [X]  Executar como Administrador pode causar problemas de permiss�es," -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]" [X]  pois os arquivos ser�o criados com o propriet�rio 'Administrador'" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]" [X]  ao inv�s do seu usu�rio normal." -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"[OK] [X] SOLU��O:" -ForegroundColor Green
 [X]  [X] Write-Host [OK]" [X]  1. Feche este PowerShell" -ForegroundColor White
 [X]  [X] Write-Host [OK]" [X]  2. Abra o PowerShell SEM privil�gios de administrador:" -ForegroundColor White
 [X]  [X] Write-Host [OK]" [X]  [X]  [X] - Pressione Win + X" -ForegroundColor Gray
 [X]  [X] Write-Host [OK]" [X]  [X]  [X] - Selecione 'Windows PowerShell' (N�O 'Windows PowerShell (Admin)')" -ForegroundColor Gray
 [X]  [X] Write-Host [OK]" [X]  3. Execute o script novamente" -ForegroundColor White
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"[i] [X] Este script N�O REQUER privil�gios de administrador." -ForegroundColor Cyan
 [X]  [X] Write-Host [OK]" [X]  Todas as opera��es s�o realizadas apenas no seu perfil de usu�rio." -ForegroundColor Cyan
 [X]  [X] Write-Host [OK]""
 [X]  [X] 
 [X]  [X] $response = Read-Host [OK]"Deseja continuar mesmo assim? (N�O recomendado) [s/N]"
 [X]  [X] if ($response -notmatch '^[Ss]$') {
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"Instala��o cancelada. Execute novamente sem privil�gios de administrador." -ForegroundColor Yellow
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] exit 0
 [X]  [X] }
 [X]  [X] 
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Warning [OK]"Continuando por solicita��o do usu�rio. Problemas de permiss�es podem ocorrer."
 [X]  [X] Write-Host [OK]""
 [X]  [X] Start-Sleep -Seconds 2
}

# Executa instala��o
try {
 [X]  [X] Install-CHAINSAWConfig
}
catch {
 [X]  [X] exit 1
}
