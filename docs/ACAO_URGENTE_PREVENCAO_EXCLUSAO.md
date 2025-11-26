# 🚨 AÇÃO URGENTE: PREVENÇÃO DE EXCLUSÃO DO PROJETO

**Data:** 26 de novembro de 2025  
**Incidentes:** 2 exclusões completas do projeto  
**Status:** CRÍTICO - MEDIDAS EMERGENCIAIS IMPLEMENTADAS

---

## 🔴 O QUE ACONTECEU

O projeto foi **completamente deletado DUAS VEZES** durante operações de limpeza.

### Incidente 1
- **Quando:** Durante limpeza de diretórios vazios
- **Causa:** Comando `Remove-Item` sem validações adequadas
- **Resultado:** Projeto inteiro deletado
- **Recuperação:** Clone do GitHub

### Incidente 2  
- **Quando:** Após implementar correções (causa ainda sob investigação)
- **Causa:** DESCONHECIDA - possivelmente edição automática ou formatação
- **Resultado:** Projeto inteiro deletado NOVAMENTE
- **Recuperação:** Clone do GitHub

---

## 🛡️ MEDIDAS EMERGENCIAIS IMPLEMENTADAS

### 1. Monitor de Integridade
**Arquivo:** `tests/Check-ProjectIntegrity.ps1`

```powershell
# Executar ANTES de qualquer operação destrutiva:
powershell -ExecutionPolicy Bypass -File .\tests\Check-ProjectIntegrity.ps1
```

**Validações:**
- ✅ Diretório do projeto existe
- ✅ `.git` está presente
- ✅ Diretórios essenciais existem
- ✅ Arquivos críticos estão presentes
- ✅ Contagem de arquivos está saudável

### 2. Sistema de Proteção
**Arquivo:** `tests/ProjectProtection.psm1`

```powershell
# Importar proteção:
Import-Module .\tests\ProjectProtection.psm1

# Usar Remove-SafeItem em vez de Remove-Item:
Remove-SafeItem -Path ".\backups" -Recurse -Force
```

**Proteções:**
- ✅ Bloqueia remoção de diretórios protegidos
- ✅ Valida presença de `.git` antes de operações
- ✅ Confirma operações recursivas grandes
- ✅ Previne exclusão do projeto root

### 3. Script de Limpeza Segura
**Arquivo:** `tests/Cleanup-EmptyDirs.ps1`

```powershell
# SEMPRE usar -WhatIf primeiro:
powershell -ExecutionPolicy Bypass -File .\tests\Cleanup-EmptyDirs.ps1 -ProjectRoot "C:\Users\csantos\chainsaw" -WhatIf

# Se OK, executar de verdade:
powershell -ExecutionPolicy Bypass -File .\tests\Cleanup-EmptyDirs.ps1 -ProjectRoot "C:\Users\csantos\chainsaw"
```

---

## ⚠️ REGRAS ABSOLUTAS

### ❌ NUNCA FAZER

1. **NUNCA** executar `Remove-Item` diretamente no projeto sem validações
2. **NUNCA** usar caminhos relativos em comandos destrutivos
3. **NUNCA** pular a validação com `-WhatIf`
4. **NUNCA** executar scripts de limpeza sem verificar integridade antes
5. **NUNCA** confiar em "Split-Path $PSScriptRoot -Parent" sem validação

### ✅ SEMPRE FAZER

1. **SEMPRE** executar `Check-ProjectIntegrity.ps1` ANTES de operações destrutivas
2. **SEMPRE** usar caminhos absolutos hardcoded
3. **SEMPRE** validar que `.git` existe antes de qualquer operação
4. **SEMPRE** usar `-WhatIf` primeiro
5. **SEMPRE** ter commit recente no GitHub antes de operações arriscadas
6. **SEMPRE** verificar integridade DEPOIS de operações

---

## 🔧 CHECKLIST PRÉ-OPERAÇÃO

Antes de executar QUALQUER comando que possa deletar arquivos:

```powershell
# 1. Verificar integridade
powershell -ExecutionPolicy Bypass -File .\tests\Check-ProjectIntegrity.ps1

# 2. Fazer commit (se houver mudanças)
git status
git add .
git commit -m "Backup antes de operacao"
git push

# 3. Executar com -WhatIf
# [seu comando aqui] -WhatIf

# 4. Se OK, executar de verdade
# [seu comando aqui]

# 5. Verificar integridade novamente
powershell -ExecutionPolicy Bypass -File .\tests\Check-ProjectIntegrity.ps1
```

---

## 🚑 RECUPERAÇÃO DE EMERGÊNCIA

Se o projeto for deletado novamente:

```powershell
# Passo 1: Ir para diretório pai
cd C:\Users\csantos

# Passo 2: Remover restos (se houver)
Remove-Item chainsaw -Recurse -Force -ErrorAction SilentlyContinue

# Passo 3: Clonar do GitHub
git clone https://github.com/chrmsantos/chainsaw.git chainsaw

# Passo 4: Verificar integridade
cd chainsaw
powershell -ExecutionPolicy Bypass -File .\tests\Check-ProjectIntegrity.ps1

# Passo 5: Confirmar Git
git status
git log --oneline -5
```

---

## 🔍 INVESTIGAÇÃO PENDENTE

### Possíveis Causas do Incidente 2

1. **Formatação Automática**
   - VS Code pode ter formatadores ativos
   - PowerShell formatting pode ter alterado scripts
   - Verificar: `.vscode/settings.json`

2. **Extensões do VS Code**
   - Alguma extensão pode estar executando scripts automaticamente
   - Verificar extensões instaladas

3. **Git Hooks**
   - Verificar se há hooks que executam scripts
   - Checar `.git/hooks/`

4. **Processos em Background**
   - Algum processo pode estar monitorando e limpando
   - Verificar Task Manager

5. **Antivírus/Segurança**
   - Windows Defender pode estar removendo arquivos
   - Verificar logs de segurança

### Ações de Investigação

```powershell
# Verificar extensões VS Code ativas
code --list-extensions

# Verificar git hooks
Get-ChildItem .git\hooks\ | Select-Object Name, LastWriteTime

# Verificar processos PowerShell
Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue

# Verificar últimas modificações
Get-ChildItem -Recurse | Sort-Object LastWriteTime -Descending | Select-Object -First 20
```

---

## 📊 ESTATÍSTICAS

- **Projeto inteiro deletado:** 2 vezes
- **Arquivos perdidos por incidente:** ~130 arquivos
- **Tempo de recuperação:** ~2 minutos (graças ao Git)
- **Commits perdidos:** 0 (tudo estava no GitHub)
- **Trabalho perdido:** Mínimo (documentação foi recriada)

---

## ✅ PRÓXIMOS PASSOS

1. [x] Monitor de integridade implementado
2. [x] Sistema de proteção criado
3. [x] Script de limpeza segura corrigido
4. [x] Documentação de emergência criada
5. [ ] Investigar causa do Incidente 2
6. [ ] Configurar alertas de integridade
7. [ ] Implementar backup automático diário
8. [ ] Revisar todas as extensões VS Code
9. [ ] Adicionar Git hooks de proteção

---

## 🎯 LIÇÃO PRINCIPAL

> **NUNCA confie em comandos destrutivos sem múltiplas camadas de validação.**

Mesmo com validações implementadas, algo pode dar errado. A única proteção real é:
1. Git com commits frequentes
2. Push regular para GitHub
3. Validação de integridade constante
4. Testes com `-WhatIf` SEMPRE

---

**Última atualização:** 26/11/2025  
**Próxima revisão:** Após cada operação de limpeza  
**Responsável:** GitHub Copilot
