# =============================================================================
# CHAINSAW - Script de Exportação de Personalizações do Word
# =============================================================================
# Versão: 1.0.0
# Licença: GNU GPLv3
# Autor: Christian Martin dos Santos
# =============================================================================

<#
.SYNOPSIS
 [X]  [X] Exporta todas as personalizações do Word do usuário atual.

.DESCRIPTION
 [X]  [X] Este script exporta:
 [X]  [X] 1. Normal.dotm (template global com macros e personalizações)
 [X]  [X] 2. Faixa de Opções Customizada (Ribbon UI)
 [X]  [X] 3. Blocos de Construção (Building Blocks)
 [X]  [X] 4. Configurações de temas e estilos
 [X]  [X] 5. Partes rápidas (Quick Parts)
 [X]  [X] 
.PARAMETER ExportPath
 [X]  [X] Caminho onde as personalizações serão exportadas.
 [X]  [X] Padrão: .\exported-config

.PARAMETER IncludeRegistry
 [X]  [X] Exporta também configurações do registro do Word.

.EXAMPLE
 [X]  [X] .\export-config.ps1
 [X]  [X] Exporta para a pasta padrão.

.EXAMPLE
 [X]  [X] .\export-config.ps1 -ExportPath [OK]"C:\Backup\WordConfig"
 [X]  [X] Exporta para caminho específico.
#>

[CmdletBinding()]
param(
 [X]  [X] [Parameter()]
 [X]  [X] [string]$ExportPath = [OK]".\exported-config",
 [X]  [X] 
 [X]  [X] [Parameter()]
 [X]  [X] [switch]$IncludeRegistry
)

$ErrorActionPreference = [OK]"Stop"

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================

$script:LogFile = $null
$script:ExportedItems = @()

# Cores
$ColorSuccess = [OK]"Green"
$ColorWarning = [OK]"Yellow"
$ColorError = [OK]"Red"
$ColorInfo = [OK]"Cyan"

# Caminhos do Word
$AppDataPath = $env:APPDATA
$LocalAppDataPath = $env:LOCALAPPDATA
$TemplatesPath = Join-Path $AppDataPath [OK]"Microsoft\Templates"
$WordSettingsPath = Join-Path $AppDataPath [OK]"Microsoft\Word"
$UiCustomizationPath = Join-Path $LocalAppDataPath [OK]"Microsoft\Office"

# =============================================================================
# FUNÇÕES DE LOG
# =============================================================================

function Initialize-LogFile {
 [X]  [X] try {
 [X]  [X]  [X]  [X] $logDir = Join-Path $ExportPath [OK]"logs"
 [X]  [X]  [X]  [X] if (-not (Test-Path $logDir)) {
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $logDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $timestamp = Get-Date -Format [OK]"yyyyMMdd_HHmmss"
 [X]  [X]  [X]  [X] $script:LogFile = Join-Path $logDir [OK]"export_$timestamp.log"
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $header = @"
================================================================================
CHAINSAW - Exportação de Personalizações do Word
================================================================================
Data/Hora: $(Get-Date -Format [OK]"dd/MM/yyyy HH:mm:ss")
Usuário: $env:USERNAME
Computador: $env:COMPUTERNAME
Sistema: $([Environment]::OSVersion.VersionString)
PowerShell: $($PSVersionTable.PSVersion)
Caminho de Exportação: $ExportPath
================================================================================

"@
 [X]  [X]  [X]  [X] Add-Content -Path $script:LogFile -Value $header
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Warning [OK]"Não foi possível criar arquivo de log: $_"
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Write-Log {
 [X]  [X] param(
 [X]  [X]  [X]  [X] [Parameter(Mandatory)]
 [X]  [X]  [X]  [X] [string]$Message,
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] [Parameter()]
 [X]  [X]  [X]  [X] [ValidateSet("INFO", [OK]"SUCCESS", [OK]"WARNING", [OK]"ERROR")]
 [X]  [X]  [X]  [X] [string]$Level = [OK]"INFO"
 [X]  [X] )
 [X]  [X] 
 [X]  [X] $timestamp = Get-Date -Format [OK]"yyyy-MM-dd HH:mm:ss"
 [X]  [X] $logEntry = [OK]"[$timestamp] [$Level] $Message"
 [X]  [X] 
 [X]  [X] if ($script:LogFile) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch { }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] switch ($Level) {
 [X]  [X]  [X]  [X] [OK]"SUCCESS" { Write-Host [OK]"[OK] $Message" -ForegroundColor $ColorSuccess }
 [X]  [X]  [X]  [X] [OK]"WARNING" { Write-Host [OK]"[!] $Message" -ForegroundColor $ColorWarning }
 [X]  [X]  [X]  [X] [OK]"ERROR" [X]  { Write-Host [OK]"[X] $Message" -ForegroundColor $ColorError }
 [X]  [X]  [X]  [X] default [X]  { Write-Host [OK]"[i] $Message" -ForegroundColor $ColorInfo }
 [X]  [X] }
}

# =============================================================================
# FUNÇÕES DE VERIFICAÇÃO
# =============================================================================

function Test-WordRunning {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Verifica se o Word está em execução.
 [X]  [X] #>
 [X]  [X] $wordProcesses = Get-Process -Name [OK]"WINWORD" -ErrorAction SilentlyContinue
 [X]  [X] return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Get-WordVersion {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Obtém a versão do Word instalada.
 [X]  [X] #>
 [X]  [X] try {
 [X]  [X]  [X]  [X] $wordPath = Get-ItemProperty [OK]"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" -ErrorAction Stop
 [X]  [X]  [X]  [X] if ($wordPath) {
 [X]  [X]  [X]  [X]  [X]  [X] $versionInfo = (Get-Item $wordPath.'(default)').VersionInfo
 [X]  [X]  [X]  [X]  [X]  [X] return $versionInfo.ProductVersion
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] catch { }
 [X]  [X] 
 [X]  [X] return $null
}

# =============================================================================
# GERENCIAMENTO DO WORD
# =============================================================================

function Test-WordRunning {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Verifica se há processos do Word em execução.
 [X]  [X] #>
 [X]  [X] $wordProcesses = Get-Process -Name [OK]"WINWORD" -ErrorAction SilentlyContinue
 [X]  [X] return ($null -ne $wordProcesses -and $wordProcesses.Count -gt 0)
}

function Stop-WordProcesses {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Fecha forçadamente todos os processos do Word.
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
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum processo do Word em execução" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] return $true
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Encontrados $($wordProcesses.Count) processo(s) do Word em execução" -Level INFO
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] foreach ($process in $wordProcesses) {
 [X]  [X]  [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Encerrando processo Word (PID: $($process.Id))..." -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if ($Force) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] # Encerra forçadamente
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $process.Kill()
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $process.WaitForExit(5000) # Aguarda até 5 segundos
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
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Ainda há $($remainingProcesses.Count) processo(s) do Word em execução" -Level WARNING
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
 [X]  [X]  [X]  [X] Solicita que o usuário salve e feche o Word, ou cancela a operação.
 [X]  [X] .DESCRIPTION
 [X]  [X]  [X]  [X] Exibe aviso ao usuário e aguarda confirmação antes de fechar o Word forçadamente.
 [X]  [X]  [X]  [X] Retorna $true se o usuário autorizar, $false se cancelar.
 [X]  [X] #>
 [X]  [X] 
 [X]  [X] # Verifica se Word está em execução
 [X]  [X] if (-not (Test-WordRunning)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Word não está em execução - prosseguindo..." -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] 
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]"¦ [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] [!] ATENÇÃO [!] [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] ¦" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"O Microsoft Word está atualmente em execução!" -ForegroundColor Yellow
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"IMPORTANTE:" -ForegroundColor Red
 [X]  [X] Write-Host [OK]" [X] • SALVE todos os seus documentos abertos no Word" -ForegroundColor White
 [X]  [X] Write-Host [OK]" [X] • FECHE o Word completamente" -ForegroundColor White
 [X]  [X] Write-Host [OK]" [X] • Outros aplicativos do Office (Excel, PowerPoint) NÃO serão afetados" -ForegroundColor Gray
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"Se você continuar, o Word será FECHADO FORÇADAMENTE e" -ForegroundColor Red
 [X]  [X] Write-Host [OK]"qualquer trabalho não salvo SERÁ PERDIDO!" -ForegroundColor Red
 [X]  [X] Write-Host [OK]""
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Word em execução - solicitando confirmação do usuário" -Level WARNING
 [X]  [X] 
 [X]  [X] # Aguarda confirmação
 [X]  [X] $response = Read-Host [OK]"Deseja FECHAR o Word e continuar a exportação? (S/N)"
 [X]  [X] 
 [X]  [X] if ($response -notmatch '^[Ss]$') {
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"? Exportação cancelada pelo usuário" -ForegroundColor Cyan
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] Salve seus documentos e execute o script novamente quando estiver pronto." -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Log [OK]"Exportação cancelada - usuário optou por não fechar o Word" -Level WARNING
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Usuário confirmou - fecha o Word
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"Fechando Microsoft Word..." -ForegroundColor Cyan
 [X]  [X] Write-Log [OK]"Usuário autorizou o fechamento do Word" -Level INFO
 [X]  [X] 
 [X]  [X] if (Stop-WordProcesses -Force) {
 [X]  [X]  [X]  [X] Write-Host [OK]"? Word fechado com sucesso" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] # Aguarda um pouco para garantir que recursos foram liberados
 [X]  [X]  [X]  [X] Start-Sleep -Seconds 2
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Host [OK]"? Não foi possível fechar o Word completamente" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Log [OK]"Falha ao fechar Word - cancelando exportação" -Level ERROR
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $retry = Read-Host [OK]"Deseja tentar novamente? (S/N)"
 [X]  [X]  [X]  [X] if ($retry -match '^[Ss]$') {
 [X]  [X]  [X]  [X]  [X]  [X] return Confirm-CloseWord # Recursão para tentar novamente
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

# =============================================================================
# FUNÇÕES DE EXPORTAÇÃO
# =============================================================================

function Export-NormalTemplate {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Exporta o template Normal.dotm.
 [X]  [X] #>
 [X]  [X] Write-Log [OK]"Exportando Normal.dotm..." -Level INFO
 [X]  [X] 
 [X]  [X] $normalPath = Join-Path $TemplatesPath [OK]"Normal.dotm"
 [X]  [X] $destPath = Join-Path $ExportPath [OK]"Templates"
 [X]  [X] 
 [X]  [X] if (-not (Test-Path $normalPath)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Normal.dotm não encontrado em: $normalPath" -Level WARNING
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] if (-not (Test-Path $destPath)) {
 [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Copy-Item -Path $normalPath -Destination $destPath -Force
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $script:ExportedItems += [PSCustomObject]@{
 [X]  [X]  [X]  [X]  [X]  [X] Type = [OK]"Normal Template"
 [X]  [X]  [X]  [X]  [X]  [X] Source = $normalPath
 [X]  [X]  [X]  [X]  [X]  [X] Destination = Join-Path $destPath [OK]"Normal.dotm"
 [X]  [X]  [X]  [X]  [X]  [X] Size = (Get-Item $normalPath).Length
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"Normal.dotm exportado com sucesso ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao exportar Normal.dotm: $_" -Level ERROR
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Export-BuildingBlocks {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Exporta os blocos de construção (Building Blocks).
 [X]  [X] #>
 [X]  [X] Write-Log [OK]"Exportando Building Blocks..." -Level INFO
 [X]  [X] 
 [X]  [X] $buildingBlocksPath = Join-Path $TemplatesPath [OK]"LiveContent\16\Managed\Word Document Building Blocks"
 [X]  [X] $userBuildingBlocksPath = Join-Path $TemplatesPath [OK]"LiveContent\16\User\Word Document Building Blocks"
 [X]  [X] $destPath = Join-Path $ExportPath [OK]"Templates\LiveContent\16"
 [X]  [X] 
 [X]  [X] $exportedCount = 0
 [X]  [X] 
 [X]  [X] # Exporta Building Blocks gerenciados (sistema)
 [X]  [X] if (Test-Path $buildingBlocksPath) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] $destManaged = Join-Path $destPath [OK]"Managed\Word Document Building Blocks"
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destManaged)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destManaged -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $buildingBlocksPath -Recurse -File
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $file.FullName.Substring($buildingBlocksPath.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destManaged $relativePath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destFile -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $exportedCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Building Blocks gerenciados: $($files.Count) arquivos" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao exportar Building Blocks gerenciados: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Exporta Building Blocks do usuário
 [X]  [X] if (Test-Path $userBuildingBlocksPath) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] $destUser = Join-Path $destPath [OK]"User\Word Document Building Blocks"
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destUser)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destUser -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $userBuildingBlocksPath -Recurse -File
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $file.FullName.Substring($userBuildingBlocksPath.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destUser $relativePath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destFile -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $exportedCount++
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $script:ExportedItems += [PSCustomObject]@{
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Type = [OK]"Building Block (User)"
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Source = $file.FullName
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Destination = $destFile
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Size = $file.Length
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Building Blocks do usuário: $($files.Count) arquivos" -Level INFO
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao exportar Building Blocks do usuário: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] if ($exportedCount -gt 0) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Building Blocks exportados: $exportedCount arquivos ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum Building Block encontrado" -Level WARNING
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Export-DocumentThemes {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Exporta temas de documentos personalizados.
 [X]  [X] #>
 [X]  [X] Write-Log [OK]"Exportando temas de documentos..." -Level INFO
 [X]  [X] 
 [X]  [X] $themesPath = Join-Path $TemplatesPath [OK]"LiveContent\16\Managed\Document Themes"
 [X]  [X] $userThemesPath = Join-Path $TemplatesPath [OK]"LiveContent\16\User\Document Themes"
 [X]  [X] $destPath = Join-Path $ExportPath [OK]"Templates\LiveContent\16"
 [X]  [X] 
 [X]  [X] $exportedCount = 0
 [X]  [X] 
 [X]  [X] # Temas gerenciados
 [X]  [X] if (Test-Path $themesPath) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] $destManaged = Join-Path $destPath [OK]"Managed\Document Themes"
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destManaged)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destManaged -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $themesPath -Recurse -File
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $file.FullName.Substring($themesPath.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destManaged $relativePath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destFile -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $exportedCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao exportar temas gerenciados: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Temas do usuário
 [X]  [X] if (Test-Path $userThemesPath) {
 [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X] $destUser = Join-Path $destPath [OK]"User\Document Themes"
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destUser)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destUser -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] $files = Get-ChildItem -Path $userThemesPath -Recurse -File
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $files) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $relativePath = $file.FullName.Substring($userThemesPath.Length + 1)
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destUser $relativePath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destDir = Split-Path $destFile -Parent
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destDir)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destDir -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $exportedCount++
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao exportar temas do usuário: $_" -Level WARNING
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] if ($exportedCount -gt 0) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Temas exportados: $exportedCount arquivos ?" -Level SUCCESS
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] else {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum tema personalizado encontrado" -Level INFO
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Export-RibbonCustomization {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Exporta personalizações da Faixa de Opções (Ribbon).
 [X]  [X] #>
 [X]  [X] Write-Log [OK]"Exportando personalização da Faixa de Opções..." -Level INFO
 [X]  [X] 
 [X]  [X] # A personalização do Ribbon é armazenada em diferentes locais dependendo da versão
 [X]  [X] $possiblePaths = @(
 [X]  [X]  [X]  [X] (Join-Path $LocalAppDataPath [OK]"Microsoft\Office\Word.officeUI"),
 [X]  [X]  [X]  [X] (Join-Path $AppDataPath [OK]"Microsoft\Office\Word.officeUI"),
 [X]  [X]  [X]  [X] (Join-Path $LocalAppDataPath [OK]"Microsoft\Office\16.0\Word.officeUI")
 [X]  [X] )
 [X]  [X] 
 [X]  [X] $destPath = Join-Path $ExportPath [OK]"RibbonCustomization"
 [X]  [X] $exportedAny = $false
 [X]  [X] 
 [X]  [X] foreach ($uiPath in $possiblePaths) {
 [X]  [X]  [X]  [X] if (Test-Path $uiPath) {
 [X]  [X]  [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destPath)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $fileName = Split-Path $uiPath -Leaf
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destPath $fileName
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $uiPath -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $script:ExportedItems += [PSCustomObject]@{
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Type = [OK]"Ribbon Customization"
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Source = $uiPath
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Destination = $destFile
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Size = (Get-Item $uiPath).Length
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Personalização do Ribbon exportada: $fileName ?" -Level SUCCESS
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $exportedAny = $true
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao exportar $uiPath : $_" -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] if (-not $exportedAny) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhuma personalização do Ribbon encontrada" -Level INFO
 [X]  [X] }
 [X]  [X] 
 [X]  [X] return $exportedAny
}

function Export-OfficeCustomUI {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Exporta arquivos de personalização da interface do Office.
 [X]  [X] #>
 [X]  [X] Write-Log [OK]"Exportando personalizações da interface..." -Level INFO
 [X]  [X] 
 [X]  [X] $customUIPath = Join-Path $LocalAppDataPath [OK]"Microsoft\Office"
 [X]  [X] $destPath = Join-Path $ExportPath [OK]"OfficeCustomUI"
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] # Procura por arquivos .officeUI
 [X]  [X]  [X]  [X] $customFiles = Get-ChildItem -Path $customUIPath -Filter [OK]"*.officeUI" -Recurse -ErrorAction SilentlyContinue
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] if ($customFiles.Count -gt 0) {
 [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destPath)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] foreach ($file in $customFiles) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destPath $file.Name
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Copy-Item -Path $file.FullName -Destination $destFile -Force
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $script:ExportedItems += [PSCustomObject]@{
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Type = [OK]"Office Custom UI"
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Source = $file.FullName
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Destination = $destFile
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Size = $file.Length
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Personalizações UI exportadas: $($customFiles.Count) arquivos ?" -Level SUCCESS
 [X]  [X]  [X]  [X]  [X]  [X] return $true
 [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] else {
 [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Nenhum arquivo de personalização UI encontrado" -Level INFO
 [X]  [X]  [X]  [X]  [X]  [X] return $false
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao exportar personalizações UI: $_" -Level WARNING
 [X]  [X]  [X]  [X] return $false
 [X]  [X] }
}

function Export-QuickAccessToolbar {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Exporta configurações da Barra de Ferramentas de Acesso Rápido.
 [X]  [X] #>
 [X]  [X] Write-Log [OK]"Exportando Barra de Ferramentas de Acesso Rápido..." -Level INFO
 [X]  [X] 
 [X]  [X] # A QAT é armazenada no arquivo .officeUI ou no registro
 [X]  [X] # Já será exportada pela função Export-OfficeCustomUI
 [X]  [X] 
 [X]  [X] Write-Log [OK]"QAT incluída nas personalizações UI" -Level INFO
 [X]  [X] return $true
}

function Export-RegistrySettings {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Exporta configurações do Word do registro.
 [X]  [X] #>
 [X]  [X] if (-not $IncludeRegistry) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Exportação do registro desabilitada (use -IncludeRegistry)" -Level INFO
 [X]  [X]  [X]  [X] return $true
 [X]  [X] }
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Exportando configurações do registro..." -Level INFO
 [X]  [X] 
 [X]  [X] $regPaths = @(
 [X]  [X]  [X]  [X] [OK]"HKCU:\Software\Microsoft\Office\16.0\Word",
 [X]  [X]  [X]  [X] [OK]"HKCU:\Software\Microsoft\Office\Common\Toolbars",
 [X]  [X]  [X]  [X] [OK]"HKCU:\Software\Microsoft\Office\16.0\Common\Toolbars"
 [X]  [X] )
 [X]  [X] 
 [X]  [X] $destPath = Join-Path $ExportPath [OK]"Registry"
 [X]  [X] $exportedAny = $false
 [X]  [X] 
 [X]  [X] foreach ($regPath in $regPaths) {
 [X]  [X]  [X]  [X] if (Test-Path $regPath) {
 [X]  [X]  [X]  [X]  [X]  [X] try {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (-not (Test-Path $destPath)) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] New-Item -Path $destPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $regFileName = $regPath -replace ':', '' -replace '\\', '_'
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $destFile = Join-Path $destPath [OK]"$regFileName.reg"
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] # Exporta a chave do registro
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $regExport = [OK]"reg export `"$regPath`" `"$destFile`" /y"
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Invoke-Expression $regExport | Out-Null
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] if (Test-Path $destFile) {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Registro exportado: $regPath ?" -Level SUCCESS
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] $exportedAny = $true
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X]  [X]  [X] catch {
 [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] Write-Log [OK]"Erro ao exportar $regPath : $_" -Level WARNING
 [X]  [X]  [X]  [X]  [X]  [X] }
 [X]  [X]  [X]  [X] }
 [X]  [X] }
 [X]  [X] 
 [X]  [X] if (-not $exportedAny) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Nenhuma configuração de registro exportada" -Level INFO
 [X]  [X] }
 [X]  [X] 
 [X]  [X] return $exportedAny
}

function Create-ExportManifest {
 [X]  [X] <#
 [X]  [X] .SYNOPSIS
 [X]  [X]  [X]  [X] Cria um manifesto com informações sobre os itens exportados.
 [X]  [X] #>
 [X]  [X] Write-Log [OK]"Criando manifesto de exportação..." -Level INFO
 [X]  [X] 
 [X]  [X] $manifest = @{
 [X]  [X]  [X]  [X] ExportDate = Get-Date -Format [OK]"yyyy-MM-dd HH:mm:ss"
 [X]  [X]  [X]  [X] UserName = $env:USERNAME
 [X]  [X]  [X]  [X] ComputerName = $env:COMPUTERNAME
 [X]  [X]  [X]  [X] WordVersion = Get-WordVersion
 [X]  [X]  [X]  [X] TotalItems = $script:ExportedItems.Count
 [X]  [X]  [X]  [X] Items = $script:ExportedItems
 [X]  [X] }
 [X]  [X] 
 [X]  [X] $manifestPath = Join-Path $ExportPath [OK]"MANIFEST.json"
 [X]  [X] $manifest | ConvertTo-Json -Depth 10 | Out-File -FilePath $manifestPath -Encoding UTF8
 [X]  [X] 
 [X]  [X] Write-Log [OK]"Manifesto criado: $manifestPath ?" -Level SUCCESS
 [X]  [X] 
 [X]  [X] # Cria também um README
 [X]  [X] $readmePath = Join-Path $ExportPath [OK]"README.txt"
 [X]  [X] $readmeContent = @"
================================================================================
CHAINSAW - Personalizações Exportadas do Word
================================================================================

Data de Exportação: $(Get-Date -Format [OK]"dd/MM/yyyy HH:mm:ss")
Usuário: $env:USERNAME
Computador: $env:COMPUTERNAME
Versão do Word: $(Get-WordVersion)

Total de itens exportados: $($script:ExportedItems.Count)

CONTEÚDO:
---------

Templates/
 [X]  [X] - Normal.dotm: Template global do Word com macros e personalizações

RibbonCustomization/
 [X]  [X] - Personalizações da Faixa de Opções (abas customizadas)

OfficeCustomUI/
 [X]  [X] - Arquivos de configuração da interface do Office

Templates/LiveContent/16/
 [X]  [X] Managed/Document Themes/
 [X]  [X]  [X]  [X] - Temas de documentos gerenciados pelo sistema
 [X]  [X] 
 [X]  [X] User/Document Themes/
 [X]  [X]  [X]  [X] - Temas personalizados pelo usuário
 [X]  [X] 
 [X]  [X] Managed/Word Document Building Blocks/
 [X]  [X]  [X]  [X] - Blocos de construção gerenciados
 [X]  [X] 
 [X]  [X] User/Word Document Building Blocks/
 [X]  [X]  [X]  [X] - Blocos de construção e partes rápidas do usuário

Registry/ (se incluído)
 [X]  [X] - Configurações do registro do Word

COMO IMPORTAR:
--------------

Para importar estas configurações em outra máquina:

1. Copie toda esta pasta para a máquina de destino

2. Execute o script de importação:
 [X]  .\import-config.ps1

Ou use o instalador principal:
 [X]  .\install.cmd

================================================================================
"@
 [X]  [X] 
 [X]  [X] $readmeContent | Out-File -FilePath $readmePath -Encoding UTF8
 [X]  [X] Write-Log [OK]"README criado: $readmePath ?" -Level SUCCESS
}

# =============================================================================
# FUNÇÃO PRINCIPAL
# =============================================================================

function Export-WordCustomizations {
 [X]  [X] Write-Host [OK]""
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Cyan
 [X]  [X] Write-Host [OK]"¦ [X]  [X]  [X]  [X] CHAINSAW - Exportação de Personalizações do Word [X]  [X]  [X]  ¦" -ForegroundColor Cyan
 [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Cyan
 [X]  [X] Write-Host [OK]""
 [X]  [X] 
 [X]  [X] # Inicializa log
 [X]  [X] Initialize-LogFile | Out-Null
 [X]  [X] Write-Log [OK]"=== INÍCIO DA EXPORTAÇÃO ===" -Level INFO
 [X]  [X] 
 [X]  [X] # Verifica e fecha Word se necessário
 [X]  [X] if (-not (Confirm-CloseWord)) {
 [X]  [X]  [X]  [X] Write-Log [OK]"Exportação cancelada - Word não foi fechado" -Level WARNING
 [X]  [X]  [X]  [X] return
 [X]  [X] }
 [X]  [X] 
 [X]  [X] # Cria pasta de exportação
 [X]  [X] if (-not (Test-Path $ExportPath)) {
 [X]  [X]  [X]  [X] New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
 [X]  [X]  [X]  [X] Write-Log [OK]"Pasta de exportação criada: $ExportPath" -Level INFO
 [X]  [X] }
 [X]  [X] 
 [X]  [X] $startTime = Get-Date
 [X]  [X] 
 [X]  [X] try {
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X] Exportando Personalizações" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]"??????????????????????????????????????????????????????????????" -ForegroundColor DarkGray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 1. Normal.dotm
 [X]  [X]  [X]  [X] Export-NormalTemplate | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 2. Building Blocks
 [X]  [X]  [X]  [X] Export-BuildingBlocks | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 3. Temas
 [X]  [X]  [X]  [X] Export-DocumentThemes | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 4. Ribbon
 [X]  [X]  [X]  [X] Export-RibbonCustomization | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 5. Custom UI
 [X]  [X]  [X]  [X] Export-OfficeCustomUI | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 6. QAT
 [X]  [X]  [X]  [X] Export-QuickAccessToolbar | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 7. Registro (opcional)
 [X]  [X]  [X]  [X] Export-RegistrySettings | Out-Null
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] # 8. Manifesto
 [X]  [X]  [X]  [X] Create-ExportManifest
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] $endTime = Get-Date
 [X]  [X]  [X]  [X] $duration = $endTime - $startTime
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]"¦ [X]  [X]  [X]  [X]  [X]  [X]  [X] EXPORTAÇÃO CONCLUÍDA COM SUCESSO! [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  ¦" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Green
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"[CHART] Resumo:" -ForegroundColor Cyan
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  • Itens exportados: $($script:ExportedItems.Count)" -ForegroundColor White
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  • Caminho: $ExportPath" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]" [X]  • Tempo decorrido: $($duration.ToString('mm\:ss'))" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"[LOG] Log: $script:LogFile" -ForegroundColor Gray
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"=== EXPORTAÇÃO CONCLUÍDA COM SUCESSO ===" -Level SUCCESS
 [X]  [X]  [X]  [X] Write-Log [OK]"Total de itens: $($script:ExportedItems.Count)" -Level INFO
 [X]  [X] }
 [X]  [X] catch {
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]"¦ [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X] ERRO NA EXPORTAÇÃO! [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  [X]  ¦" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]"+----------------------------------------------------------------+" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] Write-Host [OK]"[X] Erro: $($_.Exception.Message)" -ForegroundColor Red
 [X]  [X]  [X]  [X] Write-Host [OK]""
 [X]  [X]  [X]  [X] 
 [X]  [X]  [X]  [X] Write-Log [OK]"=== EXPORTAÇÃO FALHOU ===" -Level ERROR
 [X]  [X]  [X]  [X] Write-Log [OK]"Erro: $($_.Exception.Message)" -Level ERROR
 [X]  [X]  [X]  [X] throw
 [X]  [X] }
}

# =============================================================================
# EXECUÇÃO
# =============================================================================

Export-WordCustomizations
