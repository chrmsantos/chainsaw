# ============================================================================
# CHAINSAW - Funcoes de Backup e Atualizacao
# ============================================================================
# Versao: 3.0.0
# Licenca: GNU GPLv3
# Descricao: Sistema de backup via copy-paste e atualizacao via GitHub
# ============================================================================

function Backup-ChainsawFolder {
    <#
    .SYNOPSIS
        Gerencia backup da pasta chainsaw via copy-paste antes da instalacao.
    
    .DESCRIPTION
        1. Se existe chainsaw_old, exclui ela
        2. Copia a pasta chainsaw atual para chainsaw_old (copy-paste)
        3. Verifica integridade do backup
        
    .RETURNS
        Objeto com status e caminhos dos backups criados
    #>
    
    param(
        [Parameter()]
        [switch]$Force
    )
    
    $result = @{
        Success = $false
        BackupCreated = $false
        ChainsawPath = Join-Path $env:USERPROFILE "chainsaw"
        BackupPath = Join-Path $env:USERPROFILE "chainsaw_old"
        ErrorMessage = ""
    }
    
    try {
        Write-Host ""
        Write-Host "--------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "  ETAPA 1: Backup Automatico Pre-Instalacao (Copy-Paste)" -ForegroundColor White
        Write-Host "--------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host ""
        
        # Verificar se existe instalacao atual
        if (-not (Test-Path $result.ChainsawPath)) {
            Write-Host "[INFO] Nenhuma instalacao anterior encontrada (primeira instalacao)" -ForegroundColor Cyan
            $result.Success = $true
            return $result
        }
        
        # PASSO 1: Remover chainsaw_old se existir
        if (Test-Path $result.BackupPath) {
            Write-Host "[INFO] Removendo backup antigo: chainsaw_old" -ForegroundColor Cyan
            
            try {
                Remove-Item -Path $result.BackupPath -Recurse -Force -ErrorAction Stop
                Write-Host "[OK] Backup antigo removido" -ForegroundColor Green
            }
            catch {
                Write-Host "[AVISO] Falha ao remover backup antigo, tentando metodo alternativo..." -ForegroundColor Yellow
                
                # Metodo mais agressivo
                Get-ChildItem -Path $result.BackupPath -Recurse -Force -ErrorAction SilentlyContinue | 
                    ForEach-Object { $_.Attributes = 'Normal' }
                Start-Sleep -Milliseconds 500
                Remove-Item -Path $result.BackupPath -Recurse -Force -ErrorAction Stop
                Write-Host "[OK] Backup antigo removido (metodo alternativo)" -ForegroundColor Green
            }
        }
        
        # PASSO 2: Copiar chainsaw atual para chainsaw_old
        Write-Host "[INFO] Criando backup: chainsaw -> chainsaw_old (copy-paste)" -ForegroundColor Cyan
        
        try {
            Copy-Item -Path $result.ChainsawPath -Destination $result.BackupPath -Recurse -Force -ErrorAction Stop
            $result.BackupCreated = $true
            Write-Host "[OK] Backup criado com sucesso" -ForegroundColor Green
        }
        catch {
            $result.ErrorMessage = "Falha ao criar backup: $_"
            Write-Host "[ERRO] $($result.ErrorMessage)" -ForegroundColor Red
            return $result
        }
        
        # PASSO 3: Verificar integridade do backup
        Write-Host "[INFO] Verificando integridade do backup..." -ForegroundColor Cyan
        
        if (Test-Path $result.BackupPath) {
            $originalSize = (Get-ChildItem -Path $result.ChainsawPath -Recurse -File | Measure-Object -Property Length -Sum).Sum
            $backupSize = (Get-ChildItem -Path $result.BackupPath -Recurse -File | Measure-Object -Property Length -Sum).Sum
            
            if ($backupSize -eq $originalSize) {
                Write-Host "[OK] Integridade do backup verificada" -ForegroundColor Green
            }
            else {
                Write-Host "[AVISO] Tamanhos diferentes (Original: $originalSize bytes, Backup: $backupSize bytes)" -ForegroundColor Yellow
                Write-Host "[INFO] Isso pode ser normal devido a arquivos temporarios" -ForegroundColor Cyan
            }
        }
        else {
            $result.ErrorMessage = "Backup nao encontrado apos copia"
            Write-Host "[ERRO] $($result.ErrorMessage)" -ForegroundColor Red
            return $result
        }
        
        Write-Host ""
        Write-Host "[OK] Backup automatico concluido com sucesso" -ForegroundColor Green
        Write-Host "  - Backup criado em: chainsaw_old" -ForegroundColor Gray
        Write-Host ""
        
        $result.Success = $true
        return $result
    }
    catch {
        $result.ErrorMessage = "Erro inesperado no backup automatico: $_"
        Write-Host "[ERRO] $($result.ErrorMessage)" -ForegroundColor Red
        return $result
    }
}

function Restore-ChainsawFromBackup {
    <#
    .SYNOPSIS
        Restaura a pasta chainsaw a partir do backup chainsaw_old.
    
    .DESCRIPTION
        Remove a pasta chainsaw atual e restaura de chainsaw_old.
    #>
    
    param(
        [Parameter()]
        [switch]$Force
    )
    
    $chainsawPath = Join-Path $env:USERPROFILE "chainsaw"
    $backupPath = Join-Path $env:USERPROFILE "chainsaw_old"
    
    Write-Host ""
    Write-Host "--------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  Restauracao de Backup" -ForegroundColor White
    Write-Host "--------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    
    if (-not (Test-Path $backupPath)) {
        Write-Host "[ERRO] Nenhum backup disponivel para restaurar (chainsaw_old nao encontrado)" -ForegroundColor Red
        return $false
    }
    
    # Remove chainsaw atual se existir
    if (Test-Path $chainsawPath) {
        Write-Host "[INFO] Removendo instalacao atual..." -ForegroundColor Cyan
        try {
            Remove-Item -Path $chainsawPath -Recurse -Force -ErrorAction Stop
            Write-Host "[OK] Instalacao atual removida" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERRO] Falha ao remover instalacao atual: $_" -ForegroundColor Red
            return $false
        }
    }
    
    # Restaura do backup
    Write-Host "[INFO] Restaurando de: chainsaw_old" -ForegroundColor Cyan
    try {
        Copy-Item -Path $backupPath -Destination $chainsawPath -Recurse -Force -ErrorAction Stop
        Write-Host "[OK] Backup restaurado com sucesso" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[ERRO] Falha ao restaurar backup: $_" -ForegroundColor Red
        return $false
    }
}

function Remove-ChainsawBackups {
    <#
    .SYNOPSIS
        Remove os backups apos instalacao bem-sucedida.
    #>
    
    param(
        [Parameter()]
        [switch]$KeepBackup
    )
    
    $backupPath = Join-Path $env:USERPROFILE "chainsaw_old"
    
    Write-Host ""
    Write-Host "[INFO] Limpando backups..." -ForegroundColor Cyan
    
    if (-not $KeepBackup -and (Test-Path $backupPath)) {
        try {
            Remove-Item -Path $backupPath -Recurse -Force -ErrorAction Stop
            Write-Host "[OK] Backup removido: chainsaw_old" -ForegroundColor Green
        }
        catch {
            Write-Host "[AVISO] Nao foi possivel remover backup: $_" -ForegroundColor Yellow
        }
    }
    elseif ($KeepBackup -and (Test-Path $backupPath)) {
        Write-Host "[INFO] Backup mantido: chainsaw_old" -ForegroundColor Cyan
    }
}

function Install-ChainsawFromGitHub {
    <#
    .SYNOPSIS
        Baixa e instala a versao mais recente do CHAINSAW do GitHub.
    
    .DESCRIPTION
        1. Baixa main.zip de https://github.com/chrmsantos/chainsaw/archive/refs/heads/main.zip
        2. Descompacta o arquivo
        3. Renomeia para chainsaw_new
        4. Remove a pasta chainsaw atual
        5. Renomeia chainsaw_new para chainsaw
        6. Remove chainsaw_old se instalacao bem-sucedida
        
    .PARAMETER KeepBackup
        Mantem a pasta chainsaw_old apos instalacao bem-sucedida.
        
    .RETURNS
        $true se instalacao bem-sucedida, $false caso contrario.
    #>
    
    param(
        [Parameter()]
        [switch]$KeepBackup
    )
    
    $downloadUrl = "https://github.com/chrmsantos/chainsaw/archive/refs/heads/main.zip"
    $zipPath = Join-Path $env:TEMP "chainsaw-main.zip"
    $extractPath = Join-Path $env:TEMP "chainsaw-extract"
    $chainsawNewPath = Join-Path $env:USERPROFILE "chainsaw_new"
    $chainsawPath = Join-Path $env:USERPROFILE "chainsaw"
    $backupPath = Join-Path $env:USERPROFILE "chainsaw_old"
    
    try {
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  INSTALACAO VIA GITHUB" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        # PASSO 1: Baixar arquivo ZIP
        Write-Host "[INFO] Baixando versao mais recente do GitHub..." -ForegroundColor Cyan
        Write-Host "  URL: $downloadUrl" -ForegroundColor Gray
        
        try {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -ErrorAction Stop
            Write-Host "[OK] Download concluido" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERRO] Falha no download: $_" -ForegroundColor Red
            return $false
        }
        
        # PASSO 2: Descompactar arquivo
        Write-Host "[INFO] Descompactando arquivo..." -ForegroundColor Cyan
        
        # Limpar pasta de extracao se existir
        if (Test-Path $extractPath) {
            Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        try {
            Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force -ErrorAction Stop
            Write-Host "[OK] Arquivo descompactado" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERRO] Falha ao descompactar: $_" -ForegroundColor Red
            return $false
        }
        
        # PASSO 3: Remover arquivo ZIP
        Write-Host "[INFO] Removendo arquivo ZIP..." -ForegroundColor Cyan
        try {
            Remove-Item -Path $zipPath -Force -ErrorAction Stop
            Write-Host "[OK] Arquivo ZIP removido" -ForegroundColor Green
        }
        catch {
            Write-Host "[AVISO] Nao foi possivel remover ZIP: $_" -ForegroundColor Yellow
        }
        
        # PASSO 4: Renomear para chainsaw_new e mover para $USERPROFILE
        Write-Host "[INFO] Preparando nova instalacao..." -ForegroundColor Cyan
        
        # O GitHub cria uma pasta "chainsaw-main" dentro do ZIP
        $extractedFolder = Join-Path $extractPath "chainsaw-main"
        
        if (-not (Test-Path $extractedFolder)) {
            Write-Host "[ERRO] Pasta extraida nao encontrada: $extractedFolder" -ForegroundColor Red
            return $false
        }
        
        # Remover chainsaw_new se existir
        if (Test-Path $chainsawNewPath) {
            Remove-Item -Path $chainsawNewPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        try {
            Copy-Item -Path $extractedFolder -Destination $chainsawNewPath -Recurse -Force -ErrorAction Stop
            Write-Host "[OK] Nova instalacao preparada: chainsaw_new" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERRO] Falha ao preparar nova instalacao: $_" -ForegroundColor Red
            return $false
        }
        
        # Limpar pasta de extracao temporaria
        if (Test-Path $extractPath) {
            Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        # PASSO 5: Remover chainsaw atual
        if (Test-Path $chainsawPath) {
            Write-Host "[INFO] Removendo instalacao atual..." -ForegroundColor Cyan
            try {
                Remove-Item -Path $chainsawPath -Recurse -Force -ErrorAction Stop
                Write-Host "[OK] Instalacao atual removida" -ForegroundColor Green
            }
            catch {
                Write-Host "[ERRO] Falha ao remover instalacao atual: $_" -ForegroundColor Red
                Write-Host "[INFO] Tentando restaurar backup..." -ForegroundColor Cyan
                
                if (Restore-ChainsawFromBackup) {
                    Write-Host "[OK] Backup restaurado - sistema revertido" -ForegroundColor Green
                }
                
                return $false
            }
        }
        
        # PASSO 6: Renomear chainsaw_new para chainsaw
        Write-Host "[INFO] Instalando nova versao: chainsaw_new -> chainsaw" -ForegroundColor Cyan
        try {
            Rename-Item -Path $chainsawNewPath -NewName "chainsaw" -Force -ErrorAction Stop
            Write-Host "[OK] Nova versao instalada com sucesso" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERRO] Falha ao instalar nova versao: $_" -ForegroundColor Red
            Write-Host "[INFO] Tentando restaurar backup..." -ForegroundColor Cyan
            
            if (Restore-ChainsawFromBackup) {
                Write-Host "[OK] Backup restaurado - sistema revertido" -ForegroundColor Green
            }
            
            return $false
        }
        
        # PASSO 7: Remover backup se instalacao bem-sucedida (opcional)
        if (-not $KeepBackup -and (Test-Path $backupPath)) {
            Write-Host "[INFO] Removendo backup antigo..." -ForegroundColor Cyan
            try {
                Remove-Item -Path $backupPath -Recurse -Force -ErrorAction Stop
                Write-Host "[OK] Backup antigo removido" -ForegroundColor Green
            }
            catch {
                Write-Host "[AVISO] Nao foi possivel remover backup: $_" -ForegroundColor Yellow
                Write-Host "[INFO] Voce pode remove-lo manualmente depois: $backupPath" -ForegroundColor Cyan
            }
        }
        elseif ($KeepBackup -and (Test-Path $backupPath)) {
            Write-Host "[INFO] Backup mantido: chainsaw_old" -ForegroundColor Cyan
        }
        
        Write-Host ""
        Write-Host "[OK] Instalacao via GitHub concluida com sucesso!" -ForegroundColor Green
        Write-Host ""
        
        return $true
    }
    catch {
        Write-Host "[ERRO] Erro inesperado na instalacao: $_" -ForegroundColor Red
        
        # Tentar restaurar backup em caso de erro
        if (Test-Path $backupPath) {
            Write-Host "[INFO] Tentando restaurar backup..." -ForegroundColor Cyan
            if (Restore-ChainsawFromBackup) {
                Write-Host "[OK] Backup restaurado - sistema revertido" -ForegroundColor Green
            }
        }
        
        return $false
    }
}

function Update-ChainsawFromGitHub {
    <#
    .SYNOPSIS
        Atualiza o CHAINSAW para a versao mais recente do GitHub com backup automatico.
    
    .DESCRIPTION
        Processo completo de atualizacao:
        1. Cria backup automatico (chainsaw_old)
        2. Baixa versao mais recente do GitHub
        3. Instala nova versao
        4. Remove backup se bem-sucedido (ou mantem se especificado)
        
    .PARAMETER KeepBackup
        Mantem a pasta chainsaw_old apos atualizacao bem-sucedida.
        
    .RETURNS
        $true se atualizacao bem-sucedida, $false caso contrario.
    #>
    
    param(
        [Parameter()]
        [switch]$KeepBackup
    )
    
    Write-Host ""
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
    Write-Host "  CHAINSAW - ATUALIZACAO AUTOMATICA VIA GITHUB" -ForegroundColor White
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkCyan
    Write-Host ""
    
    # ETAPA 1: Backup
    $backupResult = Backup-ChainsawFolder
    
    if (-not $backupResult.Success) {
        Write-Host ""
        Write-Host "[ERRO] Falha no backup - atualizacao abortada" -ForegroundColor Red
        Write-Host $backupResult.ErrorMessage -ForegroundColor Red
        return $false
    }
    
    # ETAPA 2: Download e Instalacao
    $installResult = Install-ChainsawFromGitHub -KeepBackup:$KeepBackup
    
    if ($installResult) {
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
        Write-Host "  ATUALIZACAO CONCLUIDA COM SUCESSO!" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Green
        Write-Host ""
        
        if ($KeepBackup) {
            Write-Host "[INFO] Backup mantido em: $env:USERPROFILE\chainsaw_old" -ForegroundColor Cyan
            Write-Host "[INFO] Voce pode remove-lo manualmente quando desejar" -ForegroundColor Cyan
        }
        
        return $true
    }
    else {
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Red
        Write-Host "  ATUALIZACAO FALHOU" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Red
        Write-Host ""
        Write-Host "[INFO] O backup foi mantido em: $env:USERPROFILE\chainsaw_old" -ForegroundColor Cyan
        
        return $false
    }
}

function Remove-OldLogs {
    <#
    .SYNOPSIS
        Rotaciona logs mantendo apenas os 5 mais recentes.
    
    .DESCRIPTION
        Aplica politica de retencao de logs: maximo 5 arquivos por diretorio.
        Remove arquivos mais antigos baseado em LastWriteTime.
        
    .PARAMETER LogDirectory
        Caminho absoluto do diretorio de logs.
        
    .PARAMETER MaxFiles
        Numero maximo de arquivos a manter (padrao: 5).
    #>
    
    param(
        [Parameter(Mandatory)]
        [string]$LogDirectory,
        
        [Parameter()]
        [int]$MaxFiles = 5
    )
    
    if (-not (Test-Path $LogDirectory)) {
        return
    }
    
    $logFiles = Get-ChildItem -Path $LogDirectory -File -Filter "*.log" | 
        Sort-Object LastWriteTime -Descending
    
    if ($logFiles.Count -le $MaxFiles) {
        return
    }
    
    $toRemove = $logFiles | Select-Object -Skip $MaxFiles
    
    foreach ($file in $toRemove) {
        try {
            Remove-Item -Path $file.FullName -Force -ErrorAction Stop
            Write-Host "[LOG-CLEANUP] Removido: $($file.Name)" -ForegroundColor DarkGray
        }
        catch {
            Write-Host "[AVISO] Falha ao remover log: $($file.Name)" -ForegroundColor Yellow
        }
    }
}
