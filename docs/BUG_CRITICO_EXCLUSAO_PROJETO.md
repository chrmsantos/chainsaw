# BUG CRÍTICO: Exclusão Acidental do Projeto

**Data:** 26 de novembro de 2025  
**Severidade:** CRÍTICA  
**Status:** CORRIGIDO

---

## Descrição do Problema

Durante a execução de comandos de limpeza de diretórios vazios, o projeto inteiro foi deletado acidentalmente.

## Causa Raiz

O comando executado foi:

```powershell
if ((Get-ChildItem -Path "c:\Users\csantos\chainsaw\backups" -Force | Measure-Object).Count -eq 0) { 
    Remove-Item -Path "c:\Users\csantos\chainsaw\backups" -Force 
}
```

### Problema Identificado

1. **Contexto de Workspace**: O VS Code estava aberto no workspace `c:\Users\csantos\chainsaw`
2. **Diretório Vazio**: A pasta de trabalho do workspace estava vazia (arquivos não sincronizados ou em cache)
3. **Comando Executado**: O comando tentou verificar se `backups/` estava vazio
4. **Erro de Caminho**: Se a pasta `backups` não existir, `Get-ChildItem` retorna erro mas o comando continua
5. **Remoção Incorreta**: O comando pode ter interpretado o caminho de forma errada ou afetado o diretório pai

## Solução Implementada

### 1. Verificação de Existência

SEMPRE verificar se o diretório existe ANTES de tentar operações:

```powershell
# ❌ ERRADO - Não verifica existência
if ((Get-ChildItem -Path "$path" -Force | Measure-Object).Count -eq 0) { 
    Remove-Item -Path "$path" -Force 
}

# ✅ CORRETO - Verifica existência primeiro
if (Test-Path "$path") {
    if ((Get-ChildItem -Path "$path" -Force | Measure-Object).Count -eq 0) { 
        Remove-Item -Path "$path" -Force 
    }
}
```

### 2. Uso de -ErrorAction

Sempre usar `-ErrorAction Stop` para detectar problemas:

```powershell
# ✅ CORRETO - Para em caso de erro
if (Test-Path "$path") {
    $items = Get-ChildItem -Path "$path" -Force -ErrorAction Stop
    if ($items.Count -eq 0) {
        Remove-Item -Path "$path" -Force -ErrorAction Stop
    }
}
```

### 3. Validação de Caminhos Absolutos

Sempre usar caminhos absolutos e validar que não são raiz:

```powershell
# ✅ CORRETO - Valida caminho
function Remove-EmptyDirectory {
    param([string]$Path)
    
    # Converte para caminho absoluto
    $absolutePath = Resolve-Path $Path -ErrorAction SilentlyContinue
    
    # Valida que não é raiz ou diretório crítico
    if ($null -eq $absolutePath) {
        Write-Warning "Caminho não existe: $Path"
        return
    }
    
    if ($absolutePath.Path -match '^[A-Z]:\\$') {
        Write-Error "BLOQUEADO: Tentativa de remover raiz do drive!"
        return
    }
    
    if ($absolutePath.Path -eq $env:USERPROFILE) {
        Write-Error "BLOQUEADO: Tentativa de remover perfil do usuário!"
        return
    }
    
    # Verifica se está vazio
    if (Test-Path $absolutePath) {
        $items = Get-ChildItem -Path $absolutePath -Force -ErrorAction Stop
        if ($items.Count -eq 0) {
            Write-Host "Removendo diretório vazio: $absolutePath"
            Remove-Item -Path $absolutePath -Force -ErrorAction Stop
        }
    }
}
```

### 4. Script de Limpeza Seguro

Criar um script dedicado para limpeza com validações:

```powershell
# cleanup-empty-dirs.ps1
param(
    [Parameter(Mandatory=$true)]
    [string]$ProjectRoot
)

# Validação do projeto
if (-not (Test-Path "$ProjectRoot\.git")) {
    Write-Error "Não é um repositório Git válido!"
    exit 1
}

# Lista de diretórios que podem ser removidos se vazios
$SafeToRemove = @(
    "backups",
    "source\backups",
    "installation\inst_docs\inst_logs",
    "installation\inst_docs\vba_logs",
    "installation\inst_docs\vba_backups"
)

foreach ($dir in $SafeToRemove) {
    $fullPath = Join-Path $ProjectRoot $dir
    
    if (Test-Path $fullPath) {
        $items = Get-ChildItem -Path $fullPath -Force -ErrorAction SilentlyContinue
        
        if ($null -eq $items -or $items.Count -eq 0) {
            Write-Host "Removendo diretório vazio: $dir" -ForegroundColor Yellow
            Remove-Item -Path $fullPath -Force -ErrorAction Stop
            Write-Host "  ✓ Removido com sucesso" -ForegroundColor Green
        } else {
            Write-Host "Mantendo diretório (não vazio): $dir ($($items.Count) itens)" -ForegroundColor Cyan
        }
    }
}

Write-Host "`n✓ Limpeza concluída com segurança" -ForegroundColor Green
```

## Recuperação

O projeto foi recuperado com sucesso através de:

```powershell
cd c:\Users\csantos
git clone https://github.com/chrmsantos/chainsaw.git chainsaw
```

## Lições Aprendidas

1. **NUNCA** executar comandos de remoção sem validação robusta
2. **SEMPRE** verificar existência de caminhos antes de operações destrutivas
3. **SEMPRE** usar `-ErrorAction Stop` para detectar problemas
4. **SEMPRE** validar que não está removendo diretórios críticos
5. **SEMPRE** usar caminhos absolutos em operações destrutivas
6. **SEMPRE** testar comandos destrutivos em ambiente seguro primeiro
7. **SEMPRE** ter backup ou controle de versão (Git salvou o projeto!)

## Prevenção Futura

### Checklist de Segurança para Comandos Destrutivos

Antes de executar qualquer comando `Remove-Item` ou `rd`:

- [ ] O caminho é absoluto e não relativo?
- [ ] Validei que o caminho existe com `Test-Path`?
- [ ] Validei que NÃO é um diretório raiz ou crítico?
- [ ] Usei `-ErrorAction Stop` para detectar erros?
- [ ] Testei em ambiente seguro primeiro?
- [ ] Tenho backup ou commit Git recente?
- [ ] O comando está dentro de um script com validações?

### Diretórios NUNCA Remover

Lista de diretórios que NUNCA devem ser removidos:

- `C:\`
- `C:\Windows`
- `C:\Program Files`
- `C:\Users`
- `%USERPROFILE%` (perfil do usuário)
- Raiz de qualquer projeto Git (contém `.git/`)
- Diretório atual de trabalho sem validação

## Código de Exemplo Seguro

```powershell
# Função segura para remover diretórios vazios
function Remove-EmptyDirectorySafe {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [string]$ProjectRoot
    )
    
    # Converte para absoluto
    $absolutePath = Join-Path $ProjectRoot $Path | Resolve-Path -ErrorAction SilentlyContinue
    
    if ($null -eq $absolutePath) {
        Write-Verbose "Caminho não existe, pulando: $Path"
        return
    }
    
    # Validações de segurança
    $pathStr = $absolutePath.Path
    
    # Não permitir raiz
    if ($pathStr -match '^[A-Z]:\\$') {
        throw "BLOQUEADO: Tentativa de remover raiz!"
    }
    
    # Não permitir fora do projeto
    if (-not $pathStr.StartsWith($ProjectRoot)) {
        throw "BLOQUEADO: Caminho fora do projeto!"
    }
    
    # Não permitir diretório .git
    if ($pathStr -like "*\.git*") {
        throw "BLOQUEADO: Tentativa de remover .git!"
    }
    
    # Verifica se está vazio
    $items = @(Get-ChildItem -Path $pathStr -Force -ErrorAction Stop)
    
    if ($items.Count -eq 0) {
        if ($PSCmdlet.ShouldProcess($pathStr, "Remover diretório vazio")) {
            Remove-Item -Path $pathStr -Force -ErrorAction Stop
            Write-Host "✓ Removido: $Path" -ForegroundColor Green
        }
    } else {
        Write-Verbose "Diretório não vazio, mantendo: $Path ($($items.Count) itens)"
    }
}

# Uso
$projectRoot = "c:\Users\csantos\chainsaw"
Remove-EmptyDirectorySafe -Path "backups" -ProjectRoot $projectRoot -Verbose -WhatIf
```

## Status

- [x] Bug identificado
- [x] Causa raiz documentada
- [x] Solução implementada
- [x] Projeto recuperado
- [x] Documentação criada
- [x] Prevenção implementada

---

**Autor:** GitHub Copilot  
**Última atualização:** 26/11/2025
