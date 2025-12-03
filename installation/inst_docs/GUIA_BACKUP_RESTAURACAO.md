# Sistema de Backup e Restauração - CHAINSAW

## Visão Geral

O CHAINSAW implementa um sistema robusto de backup e restauração que protege as configurações do Microsoft Word antes de qualquer modificação. Este sistema garante que você possa sempre retornar ao estado anterior da instalação.

## Tipos de Backup

### 1. Backup de Templates (Automático)
**Criado automaticamente**: Durante cada instalação  
**Localização**: `%APPDATA%\Microsoft\Templates_backup_<timestamp>\`

**Conteúdo**:
- ✅ Pasta Templates completa (incluindo Normal.dotm)
- ✅ Building Blocks, temas, estilos
- ✅ Método de renomeação rápida (backup instantâneo)

**Estrutura do backup**:
```
Templates_backup_20251124_133000/
├── Normal.dotm
├── LiveContent/
│   └── 16/
│       ├── Managed/
│       └── User/
└── ... (outros arquivos da pasta Templates)
```

### 2. Backup Completo do Instalador
**Criado automaticamente**: Pelo chainsaw_installer.cmd antes da instalação  
**Localização**: `%USERPROFILE%\chainsaw_backup_<timestamp>\`

**Conteúdo**:
- ✅ Pasta Templates completa
- ✅ Personalizações do Word
- ✅ Configurações exportadas

## Scripts de Backup e Restauração

### Instalação (com backup automático)
```cmd
cd %USERPROFILE%\chainsaw\installation\inst_scripts
chainsaw_installer.cmd
```
- Cria backup completo ANTES de qualquer modificação
- Cria backup da pasta Templates (renomeação)
- Mantém os 5 backups mais recentes

### Listar Backups Disponíveis
```cmd
## Restauração de Backups

### Listar Backups Disponíveis
Para restaurar backups dos Templates criados durante a instalação:
```powershell
# Listar backups dos Templates
Get-ChildItem "$env:APPDATA\Microsoft" -Directory -Filter "Templates_backup_*" | Sort-Object Name -Descending
```

### Restaurar Backup de Templates Manualmente
Para restaurar um backup específico dos Templates:

1. Feche o Microsoft Word completamente
2. Renomeie a pasta atual:
   ```powershell
   Rename-Item "$env:APPDATA\Microsoft\Templates" "Templates_current"
   ```
3. Restaure o backup desejado:
   ```powershell
   # Exemplo: restaurar backup de 24/11/2025 às 13:30
   Rename-Item "$env:APPDATA\Microsoft\Templates_backup_20251124_133000" "Templates"
   ```

### Restaurar Backup Completo do Instalador
O backup criado pelo `chainsaw_installer.cmd` pode ser restaurado usando o script de restauração fornecido.

## Localização dos Backups

| Tipo | Caminho |
|------|---------|
| Backups do Instalador | `%USERPROFILE%\chainsaw_backup_*\` |
| Backups de Templates | `%APPDATA%\Microsoft\Templates_backup_*\` |
| Logs de Instalação | `installation\inst_docs\inst_logs\install_*.log` |

## Estrutura de Backup dos Templates

Cada backup de Templates é uma renomeação da pasta original:

```
Templates_backup_20251124_133000/
├── Normal.dotm
├── LiveContent/
│   └── 16/
│       ├── Managed/
│       │   └── Document Building Blocks.dotx
│       └── User/
│           └── Building Blocks.dotx
└── ... (outros arquivos)
      "SizeKB": 12.5
    },
```

## Processo de Instalação com Backup

```
┌─────────────────────────────────────┐
│ 1. Verificação de Pré-requisitos   │
├─────────────────────────────────────┤
│ 2. Verificação de Arquivos          │
├─────────────────────────────────────┤
│ 3. Fechar Microsoft Word            │
├─────────────────────────────────────┤
│ 4. Confirmação do Usuário           │
├─────────────────────────────────────┤
│ 5. Cópia do stamp.png               │
│    + Validação de integridade       │
├─────────────────────────────────────┤
│ 6. Backup Templates (renomeação)    │  ← BACKUP RÁPIDO
├─────────────────────────────────────┤
│ 7. Instalação Templates             │
│    + Validação de cópia             │
├─────────────────────────────────────┤
│ 8. Atualização VBA                  │
├─────────────────────────────────────┤
│ 9. Validação Final                  │  ← SEGURANÇA
│    - Verifica stamp.png             │
│    - Remove tudo se inválido        │
└─────────────────────────────────────┘
```

## Recuperação em Caso de Falha

### Recuperação Automática
Se a instalação falhar DURANTE o processo:
- A pasta `CHAINSAW` no perfil do usuário é **AUTOMATICAMENTE REMOVIDA**
- O backup da pasta Templates é restaurado automaticamente
- Mensagem de erro detalhada é exibida
- Log completo é salvo para diagnóstico

**IMPORTANTE**: Em caso de erro, o sistema garante que:
- ✅ Nenhum arquivo corrompido permanece na pasta CHAINSAW
- ✅ Nenhuma instalação parcial é deixada
- ✅ Templates são restaurados ao estado anterior

### Recuperação Manual dos Templates
Se precisar restaurar os Templates manualmente:

1. **Feche o Microsoft Word**

2. **Liste os backups disponíveis**:
   ```powershell
   Get-ChildItem "$env:APPDATA\Microsoft" -Directory -Filter "Templates_backup_*" | Sort-Object Name -Descending
   ```

3. **Restaure o backup desejado**:
   ```powershell
   # Renomeia pasta atual
   Rename-Item "$env:APPDATA\Microsoft\Templates" "Templates_old"
   
   # Restaura o backup (exemplo com timestamp 20251124_133000)
   Rename-Item "$env:APPDATA\Microsoft\Templates_backup_20251124_133000" "Templates"
   ```

## Manutenção de Backups

### Limpeza Automática
- A instalação mantém automaticamente os **5 backups mais recentes** de Templates
- Backups antigos são removidos automaticamente

### Limpeza Manual dos Templates
Para liberar espaço em disco:

```powershell
# Verificar tamanho dos backups
$templatesBackups = Get-ChildItem "$env:APPDATA\Microsoft" -Directory -Filter "Templates_backup_*"
$templatesBackups | ForEach-Object {
    $size = (Get-ChildItem $_.FullName -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
    [PSCustomObject]@{
        Nome = $_.Name
        'Tamanho (MB)' = [math]::Round($size, 2)
        Data = $_.CreationTime
    }
} | Format-Table -AutoSize

# Remover backups antigos (mais de 30 dias)
$cutoffDate = (Get-Date).AddDays(-30)
Get-ChildItem "$env:APPDATA\Microsoft" -Directory -Filter "Templates_backup_*" | 
    Where-Object { $_.CreationTime -lt $cutoffDate } | 
    Remove-Item -Recurse -Force
```

## Proteções de Segurança

### Validações Implementadas
O sistema implementa múltiplas validações para garantir segurança:

1. **Validação do stamp.png**:
   - Tamanho mínimo (> 100 bytes)
   - Integridade após cópia
   - Validação final antes de concluir

2. **Validação da instalação**:
   - Verificação completa ao final
   - Remoção automática se inválida
   - Sem arquivos corrompidos remanescentes

3. **Rollback automático**:
   - Templates restaurados em caso de erro
   - Pasta CHAINSAW removida em caso de erro
   - Estado limpo garantido

## Exemplos Práticos

### Cenário 1: Instalação com Falha
```powershell
# Se a instalação falhar, o sistema automaticamente:
# 1. Remove a pasta CHAINSAW (limpeza de segurança)
# 2. Restaura os Templates ao estado anterior
# 3. Nenhuma ação manual necessária
```

### Cenário 2: Restaurar Templates Manualmente
```powershell
# Se precisar voltar aos Templates anteriores:

# 1. Feche o Word
# 2. Liste os backups
Get-ChildItem "$env:APPDATA\Microsoft" -Directory -Filter "Templates_backup_*"

# 3. Renomeie a pasta atual
Rename-Item "$env:APPDATA\Microsoft\Templates" "Templates_old"

# 4. Restaure o backup desejado
Rename-Item "$env:APPDATA\Microsoft\Templates_backup_20251124_133000" "Templates"
```

### Cenário 3: Múltiplos Computadores
```cmd
# Exporte configurações no computador 1:
cd %USERPROFILE%\chainsaw\installation\inst_scripts
exportar_configs.cmd

# Copie a pasta exportada para o computador 2
# No computador 2, instale normalmente:
chainsaw_installer.cmd  REM Detecta e importa automaticamente
```

## Segurança e Privacidade

✅ **Todos os backups são locais** - armazenados apenas no seu computador  
✅ **Sem envio de dados** - nenhuma informação é transmitida  
✅ **Sem privilégios de admin** - funciona como usuário normal  
✅ **Código aberto** - auditável e transparente  
✅ **Backups criptografáveis** - use BitLocker ou ferramentas de criptografia se necessário  
✅ **Limpeza automática em caso de erro** - sem arquivos corrompidos  

## Perguntas Frequentes

### P: Quanto espaço os backups ocupam?
**R**: Depende do tamanho da sua pasta Templates. Normalmente entre 10-50 MB por backup.

### P: Posso mover os backups para outro local?
**R**: Sim. Os backups de Templates podem ser movidos para qualquer local seguro.

### P: O backup inclui meus documentos?
**R**: NÃO. Apenas configurações do Word são copiadas (Templates, Normal.dotm, personalizações). Seus documentos permanecem intocados.

### P: O que acontece se a instalação falhar?
**R**: O sistema automaticamente:
- Remove a pasta CHAINSAW do perfil do usuário
- Restaura os Templates ao estado anterior
- Garante que nenhum arquivo corrompido permaneça

### P: O que acontece se eu deletar os backups acidentalmente?
**R**: Não será possível restaurar para estados anteriores. Por isso, recomenda-se:
- Manter os backups em local seguro
- Fazer cópia adicional em mídia externa antes de grandes mudanças

### P: O Word precisa estar fechado para restaurar?
**R**: SIM. Sempre feche o Word antes de manipular a pasta Templates.

### P: Como sei se a instalação foi bem-sucedida?
**R**: O instalador executa uma validação final automática:
- Verifica se stamp.png foi instalado corretamente
- Remove tudo se detectar qualquer problema
- Exibe mensagem de sucesso apenas se tudo estiver correto

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
