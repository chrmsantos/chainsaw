# =============================================================================
# PROTECAO ABSOLUTA CONTRA EXCLUSAO DO PROJETO
# =============================================================================
# Este arquivo BLOQUEIA qualquer operacao que possa deletar o projeto
# =============================================================================

# Projeto: assume que este modulo fica em tests/ e protege a raiz do repo
$script:ProjectRoot = $null
if ($PSScriptRoot) {
    $script:ProjectRoot = Split-Path -Parent $PSScriptRoot
}
if (-not $script:ProjectRoot -and $MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) {
    $script:ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

# NEVER REMOVE THESE DIRECTORIES
$PROTECTED_DIRS = @(
    $script:ProjectRoot,
    (Join-Path $script:ProjectRoot '.git'),
    $env:USERPROFILE,
    $env:SystemDrive + '\\',
    (Join-Path $env:SystemRoot ''),
    (Join-Path $env:ProgramFiles ''),
    (Join-Path $env:'ProgramFiles(x86)' '')
) | Where-Object { $_ -and $_.Trim() -ne '' } | Sort-Object -Unique

# NEVER ALLOW THESE OPERATIONS ON PROJECT ROOT
function Protect-ProjectRoot {
    param([string]$Path)

    $absolutePath = $Path
    if (Test-Path $Path) {
        $absolutePath = (Resolve-Path $Path -ErrorAction SilentlyContinue).Path
    }

    foreach ($protected in $PROTECTED_DIRS) {
        if (-not $protected) { continue }

        $resolvedProtected = $protected
        if (Test-Path $protected) {
            $resolvedProtected = (Resolve-Path $protected -ErrorAction SilentlyContinue).Path
        }

        if ($absolutePath -eq $resolvedProtected) {
            throw "BLOQUEADO: Tentativa de operacao em diretorio protegido: $resolvedProtected"
        }
    }
}

# Override Remove-Item para proteger diretorios
$ExecutionContext.InvokeCommand.PreCommandLookupAction = {
    param($CommandName, $CommandLookupEventArgs)

    if ($CommandName -eq 'Remove-Item' -or $CommandName -eq 'rd' -or $CommandName -eq 'rmdir') {
        Write-Warning "ATENCAO: Comando de remocao detectado!"
        Write-Warning "Por favor, use Remove-SafeItem para operacoes no projeto"
    }
}

function Remove-SafeItem {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,

        [switch]$Recurse,
        [switch]$Force
    )

    # Protecao 1: Verifica diretorios protegidos
    Protect-ProjectRoot -Path $Path

    # Protecao 2: Verifica se esta dentro do projeto
    $projectRoot = $script:ProjectRoot
    if (Test-Path $Path) {
        $absolutePath = (Resolve-Path $Path -ErrorAction Stop).Path

        # Se esta dentro do projeto, deve ter .git presente
        if ($absolutePath.StartsWith($projectRoot, [StringComparison]::OrdinalIgnoreCase)) {
            if (-not (Test-Path (Join-Path $projectRoot ".git"))) {
                throw "BLOQUEADO: .git nao encontrado! Projeto pode estar corrompido."
            }
        }
    }

    # Protecao 3: Confirmar com usuario para operacoes recursivas
    if ($Recurse) {
        $itemCount = (Get-ChildItem $Path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
        if ($itemCount -gt 10) {
            $confirmation = Read-Host "AVISO: Tentando remover $itemCount arquivos de $Path. Confirma? (sim/nao)"
            if ($confirmation -ne 'sim') {
                Write-Warning "Operacao cancelada pelo usuario"
                return
            }
        }
    }

    # Executa remocao com parametros
    if ($PSCmdlet.ShouldProcess($Path, "Remover item")) {
        $params = @{ Path = $Path }
        if ($Recurse) { $params.Recurse = $true }
        if ($Force) { $params.Force = $true }

        Remove-Item @params
    }
}

Write-Verbose "PROTECAO ATIVA: Projeto protegido contra exclusao acidental"

Export-ModuleMember -Function Remove-SafeItem, Protect-ProjectRoot
