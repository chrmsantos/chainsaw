# Sistema de Backup e Restauração - CHAINSAW

## Visão Geral

O CHAINSAW implementa um sistema robusto de backup e restauração que protege todas as configurações do Microsoft Word antes de qualquer modificação. Este sistema garante que você possa sempre retornar ao estado anterior da instalação.

## Tipos de Backup

### 1. Backup Completo (Full Backup)
**Criado automaticamente**: Antes de cada instalação  
**Localização**: `%USERPROFILE%\CHAINSAW\backups\full_backup_<timestamp>\`

**Conteúdo**:
- ✅ Pasta Templates completa (incluindo Normal.dotm)
- ✅ Arquivo stamp.png (se existir)
- ✅ Personalizações do Word (Ribbon, Quick Access Toolbar, etc.)
- ✅ Metadados e manifesto do backup
- ✅ Arquivo README com instruções

**Estrutura do backup**:
```
full_backup_20251124_133000/
├── Templates/               # Cópia completa da pasta Templates
│   ├── Normal.dotm
│   ├── LiveContent/
│   └── ...
├── Customizations/          # Personalizações do Word
│   ├── Templates/
│   │   └── Normal.dotm
│   └── OfficeCustomUI/
│       └── *.officeUI
├── stamp.png               # Arquivo de carimbo (se existir)
├── backup_manifest.json    # Metadados do backup
└── README.txt              # Instruções de restauração
```

### 2. Backup de Templates
**Criado automaticamente**: Durante a instalação (método de renomeação rápida)  
**Localização**: `%APPDATA%\Microsoft\Templates_backup_<timestamp>\`

**Conteúdo**:
- ✅ Pasta Templates renomeada (backup instantâneo)

### 3. Backup de Personalizações
**Criado manualmente**: Ao exportar configurações  
**Localização**: `%USERPROFILE%\CHAINSAW\backups\word-customizations_<timestamp>\`

**Conteúdo**:
- ✅ Normal.dotm
- ✅ Personalizações de UI (.officeUI files)

## Scripts de Backup e Restauração

### Instalação (com backup automático)
```cmd
.\install.cmd
```
- Cria backup completo ANTES de qualquer modificação
- Cria backup da pasta Templates (renomeação)
- Mantém os 5 backups mais recentes

### Listar Backups Disponíveis
```cmd
.\restore-backup.cmd -List
```
ou
```powershell
.\restore-backup.ps1 -List
```

### Restaurar Backup Mais Recente
```cmd
.\restore-backup.cmd
```
- Mostra lista de backups
- Permite seleção interativa

### Restaurar Backup Específico
```powershell
.\restore-backup.ps1 -BackupName "20251124_133000"
```
ou
```powershell
.\restore-backup.ps1 -BackupPath "C:\Users\<user>\CHAINSAW\backups\full_backup_20251124_133000"
```

### Restaurar Apenas Templates
```powershell
.\restore-backup.ps1 -RestoreTemplates -Force
```

### Restaurar Apenas Personalizações
```powershell
.\restore-backup.ps1 -RestoreCustomizations
```

## Localização dos Backups

| Tipo | Caminho |
|------|---------|
| Backups Completos | `%USERPROFILE%\CHAINSAW\backups\full_backup_*\` |
| Backups de Templates | `%APPDATA%\Microsoft\Templates_backup_*\` |
| Backups de Personalizações | `%USERPROFILE%\CHAINSAW\backups\word-customizations_*\` |
| Logs de Instalação | `%USERPROFILE%\CHAINSAW\logs\install_*.log` |
| Logs de Restauração | `%USERPROFILE%\CHAINSAW\logs\restore_*.log` |

## Manifesto do Backup

Cada backup completo inclui um arquivo `backup_manifest.json` com metadados:

```json
{
  "Timestamp": "20251124_133000",
  "Date": "2025-11-24 13:30:00",
  "User": "csantos",
  "Computer": "WORKSTATION",
  "Items": {
    "Templates": {
      "Path": "C:\\Users\\csantos\\CHAINSAW\\backups\\full_backup_20251124_133000\\Templates",
      "SizeMB": 15.42,
      "Files": 234
    },
    "Stamp": {
      "Path": "C:\\Users\\csantos\\CHAINSAW\\backups\\full_backup_20251124_133000\\stamp.png",
      "SizeKB": 12.5
    },
    "Customizations": {
      "UIFiles": 3
    }
  }
}
```

## Processo de Instalação com Backup

```
┌─────────────────────────────────────┐
│ 1. Verificação de Pré-requisitos   │
├─────────────────────────────────────┤
│ 2. Verificação de Arquivos          │
├─────────────────────────────────────┤
│ 3. BACKUP COMPLETO ✅               │  ← PROTEÇÃO MÁXIMA
│    - Templates                      │
│    - stamp.png                      │
│    - Personalizações                │
│    - Manifesto                      │
├─────────────────────────────────────┤
│ 4. Confirmação do Usuário           │
├─────────────────────────────────────┤
│ 5. Cópia do stamp.png               │
├─────────────────────────────────────┤
│ 6. Backup Templates (renomeação)    │  ← BACKUP RÁPIDO
├─────────────────────────────────────┤
│ 7. Instalação Templates             │
├─────────────────────────────────────┤
│ 8. Atualização VBA                  │
├─────────────────────────────────────┤
│ 9. Importação Personalizações       │
└─────────────────────────────────────┘
```

## Recuperação em Caso de Falha

### Recuperação Automática
Se a instalação falhar DURANTE o processo:
- O script tenta automaticamente restaurar o backup da pasta Templates
- Mensagem de erro detalhada é exibida
- Log completo é salvo para diagnóstico

### Recuperação Manual
Se algo der errado APÓS a instalação:

1. **Liste os backups disponíveis**:
   ```cmd
   .\restore-backup.cmd -List
   ```

2. **Selecione e restaure**:
   ```cmd
   .\restore-backup.cmd
   ```
   - Digite o número do backup desejado
   - Confirme a restauração

3. **Verifique os logs**:
   ```powershell
   Get-Content "$env:USERPROFILE\CHAINSAW\logs\restore_*.log" | Select-Object -Last 50
   ```

## Manutenção de Backups

### Limpeza Automática
- A instalação mantém automaticamente os **5 backups mais recentes** de Templates
- Backups completos NÃO são removidos automaticamente (para máxima segurança)

### Limpeza Manual
Para liberar espaço em disco:

```powershell
# Verificar tamanho dos backups
$backupPath = Join-Path $env:USERPROFILE "CHAINSAW\backups"
Get-ChildItem $backupPath -Recurse | Measure-Object -Property Length -Sum

# Remover backups antigos (mais de 30 dias)
$cutoffDate = (Get-Date).AddDays(-30)
Get-ChildItem $backupPath -Directory | 
    Where-Object { $_.CreationTime -lt $cutoffDate } | 
    Remove-Item -Recurse -Force
```

## Parâmetros do Script de Restauração

| Parâmetro | Descrição |
|-----------|-----------|
| `-List` | Lista todos os backups disponíveis |
| `-BackupPath <caminho>` | Restaura de um caminho específico |
| `-BackupName <timestamp>` | Restaura pelo timestamp (ex: "20251124_133000") |
| `-Force` | Não solicita confirmação |
| `-RestoreTemplates` | Restaura apenas Templates |
| `-RestoreStamp` | Restaura apenas stamp.png |
| `-RestoreCustomizations` | Restaura apenas personalizações |

## Exemplos Práticos

### Cenário 1: Instalação com Problema
```powershell
# Você instalou e algo não funciona como esperado
# Restaure o estado anterior:
.\restore-backup.ps1
# Selecione o backup mais recente e confirme
```

### Cenário 2: Testar Nova Versão
```powershell
# Antes de testar uma nova versão, garanta que tem backup:
.\install.ps1  # Backup automático é criado

# Se não gostar, volte ao estado anterior:
.\restore-backup.ps1 -BackupName "20251124_133000" -Force
```

### Cenário 3: Múltiplos Computadores
```powershell
# Exporte configurações no computador 1:
.\export-config.ps1

# Copie a pasta exportada para o computador 2
# No computador 2, instale normalmente:
.\install.ps1  # Detecta e importa automaticamente
```

## Segurança e Privacidade

✅ **Todos os backups são locais** - armazenados apenas no seu computador  
✅ **Sem envio de dados** - nenhuma informação é transmitida  
✅ **Sem privilégios de admin** - funciona como usuário normal  
✅ **Código aberto** - auditável e transparente  
✅ **Backups criptografáveis** - use BitLocker ou ferramentas de criptografia se necessário

## Perguntas Frequentes

### P: Quanto espaço os backups ocupam?
**R**: Depende do tamanho da sua pasta Templates. Normalmente entre 10-50 MB por backup completo.

### P: Posso mover os backups para outro local?
**R**: Sim, mas será necessário especificar o caminho completo ao restaurar:
```powershell
.\restore-backup.ps1 -BackupPath "D:\Meus Backups\full_backup_20251124_133000"
```

### P: O backup inclui meus documentos?
**R**: NÃO. Apenas configurações do Word são copiadas (Templates, Normal.dotm, personalizações). Seus documentos permanecem intocados.

### P: Posso fazer backup manual a qualquer momento?
**R**: Sim, use o script de exportação:
```cmd
.\export-config.cmd
```

### P: O que acontece se eu deletar os backups acidentalmente?
**R**: Não será possível restaurar. Por isso, recomenda-se:
- Manter os backups em local seguro
- Fazer cópia adicional em mídia externa antes de grandes mudanças

### P: O Word precisa estar fechado para restaurar?
**R**: SIM. O script verifica e solicita o fechamento do Word antes de restaurar.

### P: Posso automatizar backups periódicos?
**R**: Sim, use o Agendador de Tarefas do Windows:
```powershell
# Crie uma tarefa agendada para backup semanal
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File C:\caminho\para\export-config.ps1"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 9am
Register-ScheduledTask -TaskName "CHAINSAW Backup Semanal" -Action $action -Trigger $trigger
```

## Suporte e Problemas

Se encontrar problemas:

1. **Verifique os logs**:
   ```powershell
   Get-Content "$env:USERPROFILE\CHAINSAW\logs\*.log" | Select-Object -Last 100
   ```

2. **Execute em modo verbose** (adicione `-Verbose` aos comandos PowerShell)

3. **Reporte problemas**: Inclua o log completo e descrição detalhada

## Referências

- [Guia de Instalação](GUIA_INSTALACAO.md)
- [Sem Privilégios de Admin](../docs/SEM_PRIVILEGIOS_ADMIN.md)
- [Segurança e Privacidade](../docs/SEGURANCA_PRIVACIDADE.md)
