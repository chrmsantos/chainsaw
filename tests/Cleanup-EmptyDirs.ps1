# CHAINSAW - Script de Limpeza Segura de Diretorios Vazios
# Remove apenas diretorios vazios de forma segura com multiplas validacoes

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [string]$ProjectRoot = (Split-Path $PSScriptRoot -Parent)
)

Write-Host "`n=== CHAINSAW - Limpeza Segura de Diretorios Vazios ===" -ForegroundColor Cyan
Write-Host "Diretorio do projeto: $ProjectRoot`n" -ForegroundColor Gray

# Validacao 1: Projeto existe
if (-not (Test-Path $ProjectRoot)) {
    Write-Error "ERRO: Diretorio do projeto nao existe: $ProjectRoot"
    exit 1
}

# Validacao 2: E um repositorio Git valido
if (-not (Test-Path (Join-Path $ProjectRoot ".git"))) {
    Write-Error "ERRO: Nao e um repositorio Git valido!"
    Write-Error "       .git nao encontrado em: $ProjectRoot"
    exit 1
}

# Validacao 3: Nao e raiz do sistema
if ($ProjectRoot -match '^[A-Z]:\\$') {
    Write-Error "ERRO CRITICO: Tentativa de usar raiz do drive como projeto!"
    exit 1
}

# Validacao 4: Nao e perfil do usuario
if ($ProjectRoot -eq $env:USERPROFILE) {
    Write-Error "ERRO CRITICO: Tentativa de usar perfil do usuario como projeto!"
    exit 1
}

Write-Host "OK Validacoes de seguranca passaram`n" -ForegroundColor Green

function Remove-EmptyDirectorySafe {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath,
        
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )
    
    try {
        $fullPath = Join-Path $ProjectRoot $RelativePath
        
        if (-not (Test-Path $fullPath)) {
            Write-Verbose "Diretorio nao existe, pulando: $RelativePath"
            return
        }
        
        $absolutePath = Resolve-Path $fullPath -ErrorAction Stop
        $pathStr = $absolutePath.Path
        
        # Normaliza os caminhos para comparacao (remove trailing backslash)
        $normalizedProject = $ProjectRoot.TrimEnd('\')
        $normalizedPath = $pathStr.TrimEnd('\')
        
        if (-not $normalizedPath.StartsWith($normalizedProject, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Error "BLOQUEADO: Caminho fora do projeto: $pathStr"
            return
        }
        
        if ($pathStr -like "*\.git*") {
            Write-Error "BLOQUEADO: Tentativa de remover .git: $pathStr"
            return
        }
        
        if ($pathStr -match '^[A-Z]:\\$') {
            Write-Error "BLOQUEADO: Tentativa de remover raiz: $pathStr"
            return
        }
        
        $items = @(Get-ChildItem -Path $pathStr -Force -ErrorAction Stop)
        
        if ($items.Count -eq 0) {
            if ($PSCmdlet.ShouldProcess($RelativePath, "Remover diretorio vazio")) {
                Remove-Item -Path $pathStr -Force -ErrorAction Stop
                Write-Host "  OK Removido: $RelativePath" -ForegroundColor Green
            }
        }
        else {
            $count = $items.Count
            Write-Host "  - Mantido (nao vazio): $RelativePath ($count itens)" -ForegroundColor Yellow
        }
        
    }
    catch {
        Write-Error "Erro ao processar ${RelativePath}: $_"
    }
}

$SafeDirectories = @(
    "backups",
    "source\backups",
    "installation\inst_docs\inst_logs",
    "installation\inst_docs\vba_logs",
    "installation\inst_docs\vba_backups",
    "installation\exported-config\logs",
    "tests\results"
)

Write-Host "Verificando diretorios...`n" -ForegroundColor Cyan

foreach ($dir in $SafeDirectories) {
    Write-Verbose "Processando: $dir"
    Remove-EmptyDirectorySafe -RelativePath $dir -ProjectRoot $ProjectRoot -Verbose:$VerbosePreference
}

Write-Host "`nOK Limpeza concluida com seguranca!" -ForegroundColor Green
Write-Host "  Projeto: $ProjectRoot" -ForegroundColor Gray
Write-Host "  .git preservado: " -NoNewline
if (Test-Path (Join-Path $ProjectRoot ".git")) {
    Write-Host "OK" -ForegroundColor Green
}
else {
    Write-Host "X ERRO!" -ForegroundColor Red
}

Write-Host "`n=== Estrutura do Projeto ===" -ForegroundColor Cyan
Get-ChildItem $ProjectRoot -Directory | ForEach-Object {
    $itemCount = (Get-ChildItem $_.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
    Write-Host "  $($_.Name): $itemCount arquivos" -ForegroundColor Gray
}

Write-Host ""
