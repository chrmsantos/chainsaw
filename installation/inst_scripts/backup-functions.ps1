# =============================================================================
# CHAINSAW - Funções de Backup Automático
# =============================================================================
# Versão: 2.0.3
# Licença: GNU GPLv3
# =============================================================================

function Backup-ChainsawFolder {
    <#
    .SYNOPSIS
        Gerencia backup automático da pasta chainsaw antes da instalação.
    
    .DESCRIPTION
        1. Verifica se existe chainsaw_backup em $USERPROFILE
        1.1. Se existir, renomeia para chainsaw_old_tmp_backup
        1.2. Se não existir, apenas prossegue
        2. Renomeia a pasta chainsaw atual para chainsaw_backup
        
    .RETURNS
        Objeto com status e caminhos dos backups criados
    #>
    
    param(
        [Parameter()]
        [switch]$Force
    )
    
    $result = @{
        Success = $false
        ChainsawBackupCreated = $false
        OldBackupRenamed = $false
        ChainsawPath = Join-Path $env:USERPROFILE "chainsaw"
        BackupPath = Join-Path $env:USERPROFILE "chainsaw_backup"
        OldBackupPath = Join-Path $env:USERPROFILE "chainsaw_old_tmp_backup"
        ErrorMessage = ""
    }
    
    try {
        Write-Host ""
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host "  ETAPA 1: Backup Automático Pré-Instalação" -ForegroundColor White
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
        Write-Host ""
        
        # PASSO 1: Verificar e renomear chainsaw_backup existente
        if (Test-Path $result.BackupPath) {
            Write-Host "[INFO] Backup anterior detectado: chainsaw_backup" -ForegroundColor Cyan
            
            # Remove chainsaw_old_tmp_backup se existir
            if (Test-Path $result.OldBackupPath) {
                Write-Host "[INFO] Removendo backup antigo: chainsaw_old_tmp_backup" -ForegroundColor Cyan
                
                try {
                    Remove-Item -Path $result.OldBackupPath -Recurse -Force -ErrorAction Stop
                    Write-Host "[OK] Backup antigo removido" -ForegroundColor Green
                }
                catch {
                    Write-Host "[AVISO] Não foi possível remover backup antigo: $_" -ForegroundColor Yellow
                    Write-Host "[INFO] Tentando forçar remoção..." -ForegroundColor Cyan
                    
                    # Tenta método mais agressivo
                    Get-ChildItem -Path $result.OldBackupPath -Recurse -Force -ErrorAction SilentlyContinue | 
                        ForEach-Object { $_.Attributes = 'Normal' }
                    Start-Sleep -Milliseconds 500
                    Remove-Item -Path $result.OldBackupPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Renomeia chainsaw_backup para chainsaw_old_tmp_backup
            Write-Host "[INFO] Preservando backup anterior: chainsaw_backup → chainsaw_old_tmp_backup" -ForegroundColor Cyan
            
            try {
                Rename-Item -Path $result.BackupPath -NewName "chainsaw_old_tmp_backup" -Force -ErrorAction Stop
                $result.OldBackupRenamed = $true
                Write-Host "[OK] Backup anterior preservado" -ForegroundColor Green
            }
            catch {
                $result.ErrorMessage = "Falha ao preservar backup anterior: $_"
                Write-Host "[ERRO] $($result.ErrorMessage)" -ForegroundColor Red
                return $result
            }
        }
        else {
            Write-Host "[INFO] Nenhum backup anterior encontrado" -ForegroundColor Cyan
        }
        
        # PASSO 2: Renomear chainsaw atual para chainsaw_backup
        if (Test-Path $result.ChainsawPath) {
            Write-Host "[INFO] Criando backup da instalação atual: chainsaw → chainsaw_backup" -ForegroundColor Cyan
            
            try {
                Rename-Item -Path $result.ChainsawPath -NewName "chainsaw_backup" -Force -ErrorAction Stop
                $result.ChainsawBackupCreated = $true
                Write-Host "[OK] Backup da instalação atual criado" -ForegroundColor Green
            }
            catch {
                $result.ErrorMessage = "Falha ao criar backup da instalação atual: $_"
                Write-Host "[ERRO] $($result.ErrorMessage)" -ForegroundColor Red
                
                # Se falhou, tenta restaurar o backup anterior
                if ($result.OldBackupRenamed -and (Test-Path $result.OldBackupPath)) {
                    Write-Host "[INFO] Restaurando backup anterior..." -ForegroundColor Cyan
                    try {
                        Rename-Item -Path $result.OldBackupPath -NewName "chainsaw_backup" -Force -ErrorAction SilentlyContinue
                        Write-Host "[OK] Backup anterior restaurado" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "[ERRO] Falha ao restaurar backup anterior: $_" -ForegroundColor Red
                    }
                }
                
                return $result
            }
        }
        else {
            Write-Host "[INFO] Nenhuma instalação anterior encontrada (primeira instalação)" -ForegroundColor Cyan
        }
        
        Write-Host ""
        Write-Host "[OK] Backup automático concluído com sucesso" -ForegroundColor Green
        Write-Host ""
        
        if ($result.OldBackupRenamed) {
            Write-Host "  • Backup anterior preservado: chainsaw_old_tmp_backup" -ForegroundColor Gray
        }
        if ($result.ChainsawBackupCreated) {
            Write-Host "  • Instalação atual arquivada: chainsaw_backup" -ForegroundColor Gray
        }
        Write-Host ""
        
        $result.Success = $true
        return $result
    }
    catch {
        $result.ErrorMessage = "Erro inesperado no backup automático: $_"
        Write-Host "[ERRO] $($result.ErrorMessage)" -ForegroundColor Red
        return $result
    }
}

function Restore-ChainsawFromBackup {
    <#
    .SYNOPSIS
        Restaura a pasta chainsaw a partir do backup mais recente.
    
    .DESCRIPTION
        Tenta restaurar na seguinte ordem:
        1. chainsaw_backup (backup mais recente)
        2. chainsaw_old_tmp_backup (backup anterior)
    #>
    
    param(
        [Parameter()]
        [switch]$Force
    )
    
    $chainsawPath = Join-Path $env:USERPROFILE "chainsaw"
    $backupPath = Join-Path $env:USERPROFILE "chainsaw_backup"
    $oldBackupPath = Join-Path $env:USERPROFILE "chainsaw_old_tmp_backup"
    
    Write-Host ""
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    Write-Host "  Restauração de Backup" -ForegroundColor White
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    Write-Host ""
    
    # Remove chainsaw atual se existir
    if (Test-Path $chainsawPath) {
        Write-Host "[INFO] Removendo instalação atual..." -ForegroundColor Cyan
        try {
            Remove-Item -Path $chainsawPath -Recurse -Force -ErrorAction Stop
            Write-Host "[OK] Instalação atual removida" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERRO] Falha ao remover instalação atual: $_" -ForegroundColor Red
            return $false
        }
    }
    
    # Tenta restaurar do backup mais recente
    if (Test-Path $backupPath) {
        Write-Host "[INFO] Restaurando de: chainsaw_backup" -ForegroundColor Cyan
        try {
            Rename-Item -Path $backupPath -NewName "chainsaw" -Force -ErrorAction Stop
            Write-Host "[OK] Backup restaurado com sucesso" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "[ERRO] Falha ao restaurar backup: $_" -ForegroundColor Red
        }
    }
    
    # Tenta restaurar do backup anterior
    if (Test-Path $oldBackupPath) {
        Write-Host "[INFO] Tentando backup anterior: chainsaw_old_tmp_backup" -ForegroundColor Cyan
        try {
            Rename-Item -Path $oldBackupPath -NewName "chainsaw" -Force -ErrorAction Stop
            Write-Host "[OK] Backup anterior restaurado com sucesso" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "[ERRO] Falha ao restaurar backup anterior: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "[ERRO] Nenhum backup disponível para restaurar" -ForegroundColor Red
    return $false
}

function Remove-ChainsawBackups {
    <#
    .SYNOPSIS
        Remove os backups temporários após instalação bem-sucedida.
    #>
    
    param(
        [Parameter()]
        [switch]$KeepLatest
    )
    
    $backupPath = Join-Path $env:USERPROFILE "chainsaw_backup"
    $oldBackupPath = Join-Path $env:USERPROFILE "chainsaw_old_tmp_backup"
    
    Write-Host ""
    Write-Host "[INFO] Limpando backups temporários..." -ForegroundColor Cyan
    
    # Remove backup antigo
    if (Test-Path $oldBackupPath) {
        try {
            Remove-Item -Path $oldBackupPath -Recurse -Force -ErrorAction Stop
            Write-Host "[OK] Backup antigo removido: chainsaw_old_tmp_backup" -ForegroundColor Green
        }
        catch {
            Write-Host "[AVISO] Não foi possível remover backup antigo: $_" -ForegroundColor Yellow
        }
    }
    
    # Remove backup mais recente se solicitado
    if (-not $KeepLatest -and (Test-Path $backupPath)) {
        try {
            Remove-Item -Path $backupPath -Recurse -Force -ErrorAction Stop
            Write-Host "[OK] Backup removido: chainsaw_backup" -ForegroundColor Green
        }
        catch {
            Write-Host "[AVISO] Não foi possível remover backup: $_" -ForegroundColor Yellow
        }
    }
    elseif ($KeepLatest -and (Test-Path $backupPath)) {
        Write-Host "[INFO] Backup mantido: chainsaw_backup" -ForegroundColor Cyan
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
