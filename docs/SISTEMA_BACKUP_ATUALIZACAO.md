# Sistema de Backup e Atualização - CHAINSAW

## Visão Geral

O CHAINSAW utiliza um sistema robusto de backup e atualização que garante segurança durante instalações e atualizações.

---

## Arquitetura de Backup

### Método: Copy-Paste

O sistema utiliza **copy-paste** ao invés de renomeação (move) para garantir:

- ✅ Preservação total dos dados originais
- ✅ Integridade verificável
- ✅ Recuperação confiável em caso de falha
- ✅ Sem riscos de corrupção de dados

### Estrutura de Pastas

```text
%USERPROFILE%\
├── chainsaw\          # Instalação atual
├── chainsaw_old\      # Backup da versão anterior
└── chainsaw_new\      # Nova versão (temporário durante atualização)
```

---

## Processo de Backup

### 1. Backup Pré-Instalação

Executado automaticamente antes de qualquer instalação ou atualização:

```powershell
# Função: Backup-ChainsawFolder
# Arquivo: installation/inst_scripts/backup-functions.ps1

1. Verifica se chainsaw_old existe
   └─> Se SIM: Remove chainsaw_old
2. Copia chainsaw atual para chainsaw_old (copy-paste)
3. Verifica integridade do backup (comparação de tamanhos)
```

**Resultado:**

- `chainsaw_old` contém cópia completa da instalação anterior
- Instalação original permanece intacta durante o backup

### 2. Verificação de Integridade

Após criar o backup, o sistema:

```powershell
$originalSize = (Get-ChildItem -Path $chainsawPath -Recurse -File | 
                 Measure-Object -Property Length -Sum).Sum
$backupSize = (Get-ChildItem -Path $backupPath -Recurse -File | 
               Measure-Object -Property Length -Sum).Sum

if ($backupSize -eq $originalSize) {
    # Backup OK
} else {
    # Aviso: tamanhos diferentes (pode ser normal devido a arquivos temporários)
}
```

---

## Processo de Atualização via GitHub

### Fluxo Completo

```text
┌─────────────────────────────────────────────────────────────┐
│ 1. BACKUP AUTOMÁTICO                                        │
│    chainsaw → chainsaw_old (copy-paste)                     │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ 2. DOWNLOAD DO GITHUB                                       │
│    URL: github.com/chrmsantos/chainsaw/archive/main.zip     │
│    Destino: %TEMP%\chainsaw-main.zip                        │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ 3. DESCOMPACTAÇÃO                                           │
│    chainsaw-main.zip → %TEMP%\chainsaw-extract\             │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ 4. PREPARAÇÃO                                               │
│    chainsaw-extract\chainsaw-main → chainsaw_new            │
│    Remove arquivo .zip                                      │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ 5. INSTALAÇÃO                                               │
│    Remove chainsaw atual                                    │
│    chainsaw_new → chainsaw (rename)                         │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│ 6. LIMPEZA                                                  │
│    Remove chainsaw_old (opcional, padrão: SIM)              │
│    Mantém se especificado -KeepBackup                       │
└─────────────────────────────────────────────────────────────┘
```

### Comandos

```powershell
# Atualização automática (remove backup ao final)
.\update-from-github.ps1

# Atualização mantendo backup
.\update-from-github.ps1 -KeepBackup
```

**Ou via CMD:**

```cmd
update-from-github.cmd
```

---

## Restauração de Backup

### Quando Usar

- Falha durante instalação
- Erro durante atualização
- Necessidade de reverter para versão anterior
- Problemas com nova versão

### Método Manual

```powershell
# 1. Remover instalação atual (se existir)
Remove-Item -Path "$env:USERPROFILE\chainsaw" -Recurse -Force

# 2. Copiar backup para instalação
Copy-Item -Path "$env:USERPROFILE\chainsaw_old" `
          -Destination "$env:USERPROFILE\chainsaw" `
          -Recurse -Force
```

### Método Automático

```powershell
# Carregar funções
. "$env:USERPROFILE\chainsaw\installation\inst_scripts\backup-functions.ps1"

# Executar restauração
Restore-ChainsawFromBackup
```

**Resultado:**

- Instalação atual removida
- Backup restaurado para `chainsaw`
- Sistema volta ao estado anterior

---

## Funções Disponíveis

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
    Baixa e instala versão mais recente do GitHub.
    
.PARAMETER KeepBackup
    Mantém chainsaw_old após instalação bem-sucedida.
    
.RETURNS
    $true se bem-sucedido, $false caso contrário.
#>
```

### 3. Update-ChainsawFromGitHub

```powershell
<#
.SYNOPSIS
    Atualização completa: Backup + Download + Instalação.
    
.PARAMETER KeepBackup
    Mantém chainsaw_old após atualização bem-sucedida.
    
.RETURNS
    $true se bem-sucedido, $false caso contrário.
#>
```

### 4. Restore-ChainsawFromBackup

```powershell
<#
.SYNOPSIS
    Restaura instalação a partir de chainsaw_old.
    
.RETURNS
    $true se bem-sucedido, $false caso contrário.
#>
```

### 5. Remove-ChainsawBackups

```powershell
<#
.SYNOPSIS
    Remove backups após instalação bem-sucedida.
    
.PARAMETER KeepBackup
    Se especificado, não remove o backup.
#>
```

---

## Tratamento de Erros

### Durante Backup

```text
ERRO → Interrompe processo
    └─> Instalação não inicia
    └─> Usuário decide se continua sem backup
```

### Durante Download

```text
ERRO → Restaura backup automaticamente
    └─> chainsaw_old → chainsaw
    └─> Sistema volta ao estado anterior
```

### Durante Instalação

```text
ERRO → Restaura backup automaticamente
    └─> Remove chainsaw_new (se existir)
    └─> chainsaw_old → chainsaw
    └─> Sistema volta ao estado anterior
```

---

## Segurança

### Garantias

- ✅ **Backup automático**: Sempre antes de modificações
- ✅ **Copy-paste**: Preserva dados originais durante backup
- ✅ **Verificação de integridade**: Compara tamanhos
- ✅ **Rollback automático**: Em caso de erro
- ✅ **Sem perda de dados**: Backup sempre disponível
- ✅ **Logs completos**: Rastreamento de todas operações

### Limitações

- ⚠️ Requer espaço em disco para cópia completa (~20-50 MB)
- ⚠️ Requer conexão com Internet (para download)
- ⚠️ Verificação de integridade baseada em tamanho (não hash)

---

## Exemplos de Uso

### Atualização Simples

```powershell
cd "$env:USERPROFILE\chainsaw\installation\inst_scripts"
.\update-from-github.ps1
```

**Saída:**

```text
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  CHAINSAW - ATUALIZAÇÃO AUTOMÁTICA VIA GITHUB
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

--------------------------------------------------------------
  ETAPA 1: Backup Automático Pré-Instalação (Copy-Paste)
--------------------------------------------------------------

[INFO] Removendo backup antigo: chainsaw_old
[OK] Backup antigo removido
[INFO] Criando backup: chainsaw -> chainsaw_old (copy-paste)
[OK] Backup criado com sucesso
[INFO] Verificando integridade do backup...
[OK] Integridade do backup verificada

[OK] Backup automático concluído com sucesso
  - Backup criado em: chainsaw_old

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  INSTALAÇÃO VIA GITHUB
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[INFO] Baixando versão mais recente do GitHub...
  URL: https://github.com/chrmsantos/chainsaw/archive/refs/heads/main.zip
[OK] Download concluído
[INFO] Descompactando arquivo...
[OK] Arquivo descompactado
[INFO] Removendo arquivo ZIP...
[OK] Arquivo ZIP removido
[INFO] Preparando nova instalação...
[OK] Nova instalação preparada: chainsaw_new
[INFO] Removendo instalação atual...
[OK] Instalação atual removida
[INFO] Instalando nova versão: chainsaw_new -> chainsaw
[OK] Nova versão instalada com sucesso
[INFO] Removendo backup antigo...
[OK] Backup antigo removido

[OK] Instalação via GitHub concluída com sucesso!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ATUALIZAÇÃO CONCLUÍDA COM SUCESSO!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

### Atualização Mantendo Backup

```powershell
.\update-from-github.ps1 -KeepBackup
```

### Restauração Manual

```powershell
. .\backup-functions.ps1
Restore-ChainsawFromBackup
```

---

## Manutenção

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

**Versão:** 3.0.0  
**Última atualização:** 3 de dezembro de 2024
