## 🔥 INCIDENTE CRÍTICO: EXCLUSÃO REPETIDA DO PROJETO

**Data:** 26 de novembro de 2025  
**Frequência:** 3 EXCLUSÕES EM POUCAS HORAS  
**Severidade:** CRÍTICA - PERDA DE PRODUTIVIDADE TOTAL

---

## 📊 REGISTRO DE INCIDENTES

| # | Horário (aprox) | Ação Tomada | Status |
|---|----------------|-------------|--------|
| 1 | ~13:00 | Exclusão durante limpeza manual | ✅ Recuperado |
| 2 | ~13:30 | Exclusão causa desconhecida | ✅ Recuperado |
| 3 | ~14:00 | Exclusão causa desconhecida | ✅ Recuperado |

---

## 🔍 HIPÓTESES DA CAUSA

### 1. ⚠️ Formatação Automática de Scripts

- **Probabilidade:** ALTA
- **Evidência:** "Some edits were made, by the user or possibly by a formatter"
- **Scripts afetados:** `Check-ProjectIntegrity.ps1`, `Cleanup-EmptyDirs.ps1`
- **Suspeito:** VS Code PowerShell Extension com formatação automática

### 2. ⚠️ Extensão VS Code Executando Scripts

- **Probabilidade:** MÉDIA
- **Evidência:** Scripts sendo modificados automaticamente
- **Suspeito:** PowerShell Extension com "Run on Save" ativado

### 3. ⚠️ Git Hooks ou Automação

- **Probabilidade:** BAIXA
- **Evidência:** Nenhuma até o momento

### 4. ⚠️ Antivírus/Windows Defender

- **Probabilidade:** BAIXA
- **Evidência:** Nenhuma até o momento

---

## 🛡️ MEDIDAS EMERGENCIAIS IMPLEMENTADAS

### Proteção em Múltiplas Camadas

1. **Backup Automático Permanente**
   - Localização: `C:\Users\csantos\chainsaw_backup_permanente\`
   - Script: `C:\Users\csantos\PROTECAO_CHAINSAW.ps1`
   - Mantém últimos 5 backups incrementais

2. **Recuperação Automática**
   - Detecta ausência do projeto
   - Restaura do backup local primeiro
   - Fallback para clone do GitHub

3. **Monitor de Integridade**
   - `tests/Check-ProjectIntegrity.ps1`
   - Valida .git, diretórios e arquivos essenciais

---

## ⚡ AÇÃO IMEDIATA NECESSÁRIA

### DESATIVAR FORMATAÇÃO AUTOMÁTICA

Editar `.vscode/settings.json`:

```json
{
  "powershell.scriptAnalysis.enable": false,
  "editor.formatOnSave": false,
  "files.autoSave": "off",
  "[powershell]": {
    "editor.formatOnSave": false,
    "editor.formatOnPaste": false,
    "editor.formatOnType": false
  }
}
```

### VERIFICAR EXTENSÕES VS CODE

```powershell
code --list-extensions | Out-File C:\Users\csantos\vscode_extensions.txt
```

Extensões suspeitas:

- `ms-vscode.powershell` - PowerShell Extension
- Qualquer extensão de formatação automática
- Extensões de limpeza de arquivos

### DESATIVAR RUN ON SAVE

Verificar se há configurações que executam scripts ao salvar:

- PowerShell Extension settings
- Task Runner extensions
- File Watcher extensions

---

## 🚫 REGRAS ABSOLUTAS (ATUALIZADAS)

### ❌ NUNCA MAIS FAZER

1. **NUNCA** confiar em formatadores automáticos
2. **NUNCA** permitir "Format on Save" em scripts de sistema
3. **NUNCA** deixar scripts de limpeza executáveis sem supervisão
4. **NUNCA** criar scripts que usam `$PSScriptRoot` sem validação absoluta

### ✅ SEMPRE FAZER

1. **SEMPRE** manter backup permanente atualizado
2. **SEMPRE** desabilitar formatação em scripts críticos
3. **SEMPRE** verificar mudanças antes de salvar
4. **SEMPRE** executar `PROTECAO_CHAINSAW.ps1` regularmente

---

## 📋 INVESTIGAÇÃO PENDENTE

### Próximos Passos de Diagnóstico

```powershell
# 1. Verificar configurações VS Code
Get-Content .vscode\settings.json

# 2. Listar extensões ativas
code --list-extensions

# 3. Verificar processos PowerShell em execução
Get-Process -Name powershell, pwsh -IncludeUserName

# 4. Verificar histórico de arquivos modificados
Get-ChildItem -Recurse | 
  Sort-Object LastWriteTime -Descending | 
  Select-Object -First 20 FullName, LastWriteTime

# 5. Verificar logs do Windows
Get-EventLog -LogName System -Newest 50 | 
  Where-Object { $_.Message -like "*chainsaw*" }
```

---

## 🔧 CONFIGURAÇÃO SEGURA DO VS CODE

Criar/atualizar `.vscode/settings.json`:

```json
{
  "powershell.scriptAnalysis.enable": false,
  "powershell.codeFormatting.autoCorrectAliases": false,
  "powershell.codeFormatting.useCorrectCasing": false,
  "powershell.integratedConsole.focusConsoleOnExecute": false,
  "editor.formatOnSave": false,
  "editor.formatOnPaste": false,
  "editor.formatOnType": false,
  "files.autoSave": "off",
  "files.trimTrailingWhitespace": false,
  "[powershell]": {
    "editor.formatOnSave": false,
    "editor.formatOnPaste": false,
    "editor.formatOnType": false,
    "editor.defaultFormatter": null
  },
  "task.autoDetect": "off",
  "npm.autoDetect": "off"
}
```

---

## 🎯 SOLUÇÃO DEFINITIVA

### Opção 1: Mover Projeto Para Fora do Workspace VS Code

```powershell
# Mover para local seguro
$safeLocation = "C:\Projects\chainsaw"
Move-Item C:\Users\csantos\chainsaw $safeLocation
# Abrir VS Code sem workspace
```

### Opção 2: Usar .gitignore para Proteger Scripts

Adicionar ao `.gitignore`:

```
# Proteger scripts de formatação
tests/*EmptyDirs.ps1
```

### Opção 3: Desabilitar Completamente PowerShell Extension

```powershell
code --disable-extension ms-vscode.powershell
```

---

## 📊 IMPACTO

- **Tempo perdido:** ~1-2 horas
- **Produtividade:** 0% durante incidentes
- **Risco de perda de trabalho:** ALTO (se não comitado)
- **Confiança no ambiente:** BAIXA

---

## ✅ STATUS ATUAL

- [x] Projeto recuperado (3ª vez)
- [x] Backup permanente criado
- [x] Script de proteção implementado
- [x] Documentação de emergência atualizada
- [ ] **URGENTE:** Identificar e desabilitar causa raiz
- [ ] Configurar VS Code de forma segura
- [ ] Mover projeto para local seguro (se necessário)

---

## 🚨 PRÓXIMA AÇÃO

**IMEDIATA:** Desabilitar formatação automática no VS Code

**CRÍTICA:** Identificar qual processo está deletando os arquivos

**PREVENTIVA:** Executar `PROTECAO_CHAINSAW.ps1` a cada 15 minutos

---

**Última atualização:** 26/11/2025 - Após 3º incidente  
**Próxima revisão:** A cada incidente (esperamos que não haja mais!)  
**Status:** 🔴 CRÍTICO - INVESTIGAÇÃO EM ANDAMENTO
