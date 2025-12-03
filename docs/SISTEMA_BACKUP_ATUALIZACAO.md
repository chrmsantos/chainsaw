# Sistema de Backup e AtualizaÃ§Ã£o - CHAINSAW

## VisÃ£o Geral

O CHAINSAW utiliza um sistema robusto de backup e atualizaÃ§Ã£o que garante seguranÃ§a durante instalaÃ§Ãµes e atualizaÃ§Ãµes.

---

## Arquitetura de Backup

### MÃ©todo: Copy-Paste

O sistema utiliza **copy-paste** ao invÃ©s de renomeaÃ§Ã£o (move) para garantir:

- âœ… PreservaÃ§Ã£o total dos dados originais
- âœ… Integridade verificÃ¡vel
- âœ… RecuperaÃ§Ã£o confiÃ¡vel em caso de falha
- âœ… Sem riscos de corrupÃ§Ã£o de dados

### Estrutura de Pastas

```text
%USERPROFILE%\
â”œâ”€â”€ chainsaw\          # InstalaÃ§Ã£o atual
â”œâ”€â”€ chainsaw_old\      # Backup da versÃ£o anterior
â””â”€â”€ chainsaw_new\      # Nova versÃ£o (temporÃ¡rio durante atualizaÃ§Ã£o)

```

---

## Processo de Backup

### 1. Backup PrÃ©-InstalaÃ§Ã£o

Executado automaticamente antes de qualquer instalaÃ§Ã£o ou atualizaÃ§Ã£o:

```powershell
# FunÃ§Ã£o: Backup-ChainsawFolder
# Arquivo: installation/inst_scripts/backup-functions.ps1

1. Verifica se chainsaw_old existe
   â””â”€> Se SIM: Remove chainsaw_old
2. Copia chainsaw atual para chainsaw_old (copy-paste)
3. Verifica integridade do backup (comparaÃ§Ã£o de tamanhos)

```

**Resultado:**

- `chainsaw_old` contÃ©m cÃ³pia completa da instalaÃ§Ã£o anterior
- InstalaÃ§Ã£o original permanece intacta durante o backup

### 2. VerificaÃ§Ã£o de Integridade

ApÃ³s criar o backup, o sistema:

```powershell
$originalSize = (Get-ChildItem -Path $chainsawPath -Recurse -File | 
                 Measure-Object -Property Length -Sum).Sum
$backupSize = (Get-ChildItem -Path $backupPath -Recurse -File | 
               Measure-Object -Property Length -Sum).Sum

if ($backupSize -eq $originalSize) {
    # Backup OK
} else {
    # Aviso: tamanhos diferentes (pode ser normal devido a arquivos temporÃ¡rios)
}

```

---

## Processo de AtualizaÃ§Ã£o via GitHub

### Fluxo Completo

```text
â”Œâ”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”
â”‚ 1. BACKUP AUTOMÃTICO                                        â”‚
â”‚    chainsaw â†’ chainsaw_old (copy-paste)                     â”‚
â””â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”˜
                            â†“
â”Œâ”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”
â”‚ 2. DOWNLOAD DO GITHUB                                       â”‚
â”‚    URL: github.com/chrmsantos/chainsaw/archive/main.zip     â”‚
â”‚    Destino: %TEMP%\chainsaw-main.zip                        â”‚
â””â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”˜
                            â†“
â”Œâ”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”
â”‚ 3. DESCOMPACTAÃ‡ÃƒO                                           â”‚
â”‚    chainsaw-main.zip â†’ %TEMP%\chainsaw-extract\             â”‚
â””â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”˜
                            â†“
â”Œâ”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”
â”‚ 4. PREPARAÃ‡ÃƒO                                               â”‚
â”‚    chainsaw-extract\chainsaw-main â†’ chainsaw_new            â”‚
â”‚    Remove arquivo .zip                                      â”‚
â””â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”˜
                            â†“
â”Œâ”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”
â”‚ 5. INSTALAÃ‡ÃƒO                                               â”‚
â”‚    Remove chainsaw atual                                    â”‚
â”‚    chainsaw_new â†’ chainsaw (rename)                         â”‚
â””â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”˜
                            â†“
â”Œâ”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”
â”‚ 6. LIMPEZA                                                  â”‚
â”‚    Remove chainsaw_old (opcional, padrÃ£o: SIM)              â”‚
â”‚    MantÃ©m se especificado -KeepBackup                       â”‚
â””â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”˜

```

### Comandos

```powershell
# AtualizaÃ§Ã£o automÃ¡tica (remove backup ao final)
.\update-from-github.ps1

# AtualizaÃ§Ã£o mantendo backup
.\update-from-github.ps1 -KeepBackup

```

**Ou via CMD:**

```cmd
update-from-github.cmd

```

---

## RestauraÃ§Ã£o de Backup

### Quando Usar

- Falha durante instalaÃ§Ã£o
- Erro durante atualizaÃ§Ã£o
- Necessidade de reverter para versÃ£o anterior
- Problemas com nova versÃ£o

### MÃ©todo Manual

```powershell
# 1. Remover instalaÃ§Ã£o atual (se existir)
Remove-Item -Path "$env:USERPROFILE\chainsaw" -Recurse -Force

# 2. Copiar backup para instalaÃ§Ã£o
Copy-Item -Path "$env:USERPROFILE\chainsaw_old" `
          -Destination "$env:USERPROFILE\chainsaw" `
          -Recurse -Force

```

### MÃ©todo AutomÃ¡tico

```powershell
# Carregar funÃ§Ãµes
. "$env:USERPROFILE\chainsaw\installation\inst_scripts\backup-functions.ps1"

# Executar restauraÃ§Ã£o
Restore-ChainsawFromBackup

```

**Resultado:**

- InstalaÃ§Ã£o atual removida
- Backup restaurado para `chainsaw`
- Sistema volta ao estado anterior

---

## FunÃ§Ãµes DisponÃ­veis

### 1. Backup-ChainsawFolder

```powershell
<#
.SYNOPSIS
    Cria backup da pasta chainsaw via copy-paste.
    
.RETURNS
    @{
        Success = $true/$false
        BackupCreated = $true/$false
        ChainsawPath = "C:\Users\...\chainsaw"
        BackupPath = "C:\Users\...\chainsaw_old"
        ErrorMessage = "..."
    }
#>

```

### 2. Install-ChainsawFromGitHub

```powershell
<#
.SYNOPSIS
    Baixa e instala versÃ£o mais recente do GitHub.
    
.PARAMETER KeepBackup
    MantÃ©m chainsaw_old apÃ³s instalaÃ§Ã£o bem-sucedida.
    
.RETURNS
    $true se bem-sucedido, $false caso contrÃ¡rio.
#>

```

### 3. Update-ChainsawFromGitHub

```powershell
<#
.SYNOPSIS
    AtualizaÃ§Ã£o completa: Backup + Download + InstalaÃ§Ã£o.
    
.PARAMETER KeepBackup
    MantÃ©m chainsaw_old apÃ³s atualizaÃ§Ã£o bem-sucedida.
    
.RETURNS
    $true se bem-sucedido, $false caso contrÃ¡rio.
#>

```

### 4. Restore-ChainsawFromBackup

```powershell
<#
.SYNOPSIS
    Restaura instalaÃ§Ã£o a partir de chainsaw_old.
    
.RETURNS
    $true se bem-sucedido, $false caso contrÃ¡rio.
#>

```

### 5. Remove-ChainsawBackups

```powershell
<#
.SYNOPSIS
    Remove backups apÃ³s instalaÃ§Ã£o bem-sucedida.
    
.PARAMETER KeepBackup
    Se especificado, nÃ£o remove o backup.
#>

```

---

## Tratamento de Erros

### Durante Backup

```text
ERRO â†’ Interrompe processo
    â””â”€> InstalaÃ§Ã£o nÃ£o inicia
    â””â”€> UsuÃ¡rio decide se continua sem backup

```

### Durante Download

```text
ERRO â†’ Restaura backup automaticamente
    â””â”€> chainsaw_old â†’ chainsaw
    â””â”€> Sistema volta ao estado anterior

```

### Durante InstalaÃ§Ã£o

```text
ERRO â†’ Restaura backup automaticamente
    â””â”€> Remove chainsaw_new (se existir)
    â””â”€> chainsaw_old â†’ chainsaw
    â””â”€> Sistema volta ao estado anterior

```

---

## SeguranÃ§a

### Garantias

- âœ… **Backup automÃ¡tico**: Sempre antes de modificaÃ§Ãµes
- âœ… **Copy-paste**: Preserva dados originais durante backup
- âœ… **VerificaÃ§Ã£o de integridade**: Compara tamanhos
- âœ… **Rollback automÃ¡tico**: Em caso de erro
- âœ… **Sem perda de dados**: Backup sempre disponÃ­vel
- âœ… **Logs completos**: Rastreamento de todas operaÃ§Ãµes

### LimitaÃ§Ãµes

- âš ï¸ Requer espaÃ§o em disco para cÃ³pia completa (~20-50 MB)
- âš ï¸ Requer conexÃ£o com Internet (para download)
- âš ï¸ VerificaÃ§Ã£o de integridade baseada em tamanho (nÃ£o hash)

---

## Exemplos de Uso

### AtualizaÃ§Ã£o Simples

```powershell
cd "$env:USERPROFILE\chainsaw\installation\inst_scripts"
.\update-from-github.ps1

```

**SaÃ­da:**

```text
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
  CHAINSAW - ATUALIZAÃ‡ÃƒO AUTOMÃTICA VIA GITHUB
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

--------------------------------------------------------------
  ETAPA 1: Backup AutomÃ¡tico PrÃ©-InstalaÃ§Ã£o (Copy-Paste)
--------------------------------------------------------------

[INFO] Removendo backup antigo: chainsaw_old
[OK] Backup antigo removido
[INFO] Criando backup: chainsaw -> chainsaw_old (copy-paste)
[OK] Backup criado com sucesso
[INFO] Verificando integridade do backup...
[OK] Integridade do backup verificada

[OK] Backup automÃ¡tico concluÃ­do com sucesso
  - Backup criado em: chainsaw_old

â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
  INSTALAÃ‡ÃƒO VIA GITHUB
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

[INFO] Baixando versÃ£o mais recente do GitHub...
  URL: https://github.com/chrmsantos/chainsaw/archive/refs/heads/main.zip
[OK] Download concluÃ­do
[INFO] Descompactando arquivo...
[OK] Arquivo descompactado
[INFO] Removendo arquivo ZIP...
[OK] Arquivo ZIP removido
[INFO] Preparando nova instalaÃ§Ã£o...
[OK] Nova instalaÃ§Ã£o preparada: chainsaw_new
[INFO] Removendo instalaÃ§Ã£o atual...
[OK] InstalaÃ§Ã£o atual removida
[INFO] Instalando nova versÃ£o: chainsaw_new -> chainsaw
[OK] Nova versÃ£o instalada com sucesso
[INFO] Removendo backup antigo...
[OK] Backup antigo removido

[OK] InstalaÃ§Ã£o via GitHub concluÃ­da com sucesso!

â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
  ATUALIZAÃ‡ÃƒO CONCLUÃDA COM SUCESSO!
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

```

### AtualizaÃ§Ã£o Mantendo Backup

```powershell
.\update-from-github.ps1 -KeepBackup

```

### RestauraÃ§Ã£o Manual

```powershell
. .\backup-functions.ps1
Restore-ChainsawFromBackup

```

---

## ManutenÃ§Ã£o

### Remover Backup Manualmente

```powershell
Remove-Item -Path "$env:USERPROFILE\chainsaw_old" -Recurse -Force

```

### Verificar Tamanho do Backup

```powershell
$size = (Get-ChildItem -Path "$env:USERPROFILE\chainsaw_old" -Recurse -File | 
         Measure-Object -Property Length -Sum).Sum / 1MB
Write-Host "Tamanho do backup: $([Math]::Round($size, 2)) MB"

```

---

**VersÃ£o:** 3.0.0  
**Ãšltima atualizaÃ§Ã£o:** 3 de dezembro de 2024
