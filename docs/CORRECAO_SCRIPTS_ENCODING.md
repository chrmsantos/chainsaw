# Correção dos Scripts Corrompidos por Encoding

## Data
07/11/2024

## Problema Identificado

Os scripts PowerShell (`install.ps1` e `export-config.ps1`) foram corrompidos após tentativas de correção de encoding nos commits:
- **31efecc** - "CORRIGIDO ENCODING ASCII / UTF-8"
- **2ed7d37** - "CORRIGIDO ENCODING ASCII / UTF-8"

### Sintomas

Os arquivos `.ps1` continham caracteres inválidos como:
- `[X]` substituindo espaços em branco
- `[OK]` substituindo aspas duplas (`"`)
- Outros caracteres especiais corrompidos

Exemplo do código corrompido:
```powershell
 [X]  [X] [string]$SourcePath = [OK]"",
```

Ao invés de:
```powershell
    [string]$SourcePath = "",
```

### Causa Raiz

As tentativas de "corrigir" o encoding dos arquivos acabaram introduzindo uma conversão incorreta de caracteres, substituindo:
- Espaços por `[X]`
- Aspas duplas por `[OK]`
- Outros caracteres especiais foram mal interpretados

Isso tornou os scripts **completamente não funcionais** e impossíveis de executar.

## Solução Aplicada

### 1. Identificação da Versão Funcional

Usando o histórico do git, identificamos que o commit **210f223** ("ok") continha a última versão funcional dos scripts antes das correções de encoding problemáticas.

```powershell
git log --oneline -10
```

### 2. Restauração dos Arquivos

Restauramos os arquivos para a versão funcional:

```powershell
git checkout 210f223 -- install.ps1 export-config.ps1
```

### 3. Verificação

Testamos os scripts restaurados:

```powershell
# Verificar se não há mais caracteres corrompidos
Get-Content install.ps1 | Select-Object -First 20 | ForEach-Object { 
    if ($_ -match '\[X\]|\[OK\]') { 
        Write-Host "ERRO: Arquivo ainda corrompido" -ForegroundColor Red 
    } 
}

# Testar carregamento do script
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& { 
    $ErrorActionPreference = 'Stop'; 
    . .\install.ps1 -WhatIf 2>&1 
}"

# Verificar documentação
Get-Help .\install.ps1 -Detailed
Get-Help .\export-config.ps1 -Detailed
```

### 4. Commit e Push

Criamos um commit documentando a correção:

```
REVERTIDO: Restaurados scripts corrompidos pelas correções de encoding

Os commits anteriores (31efecc e 2ed7d37) tentaram corrigir encoding mas introduziram
caracteres inválidos ([X], [OK]) que corromperam os arquivos .ps1.

Restaurados os arquivos para a versão funcional do commit 210f223.
```

## Arquivos Corrigidos

- ✅ `install.ps1` - Restaurado e funcionando
- ✅ `export-config.ps1` - Restaurado e funcionando

## Status Atual

### ✅ Scripts Funcionais

Ambos os scripts agora:
- Carregam sem erros
- Exibem ajuda corretamente com `Get-Help`
- Podem ser executados via `.cmd` launchers
- Têm toda a documentação intacta

### Encoding Correto

Os arquivos estão salvos com encoding apropriado para PowerShell (UTF-8 com BOM), mantendo:
- Caracteres acentuados corretos (ç, ã, é, etc.)
- Aspas e espaços corretos
- Comentários formatados adequadamente

## Lições Aprendidas

### ⚠️ NÃO FAZER

1. **Nunca** tentar "corrigir" encoding manualmente sem entender o problema
2. **Nunca** usar ferramentas de conversão automática sem validar o resultado
3. **Nunca** fazer commit de arquivos sem testar se funcionam

### ✅ FAZER

1. **Sempre** testar scripts após qualquer modificação
2. **Sempre** verificar o diff antes de fazer commit
3. **Sempre** manter backups ou confiar no histórico do git
4. Se houver problemas de encoding:
   - Identificar qual é o encoding atual
   - Identificar qual deveria ser o encoding correto
   - Usar ferramentas apropriadas (VS Code, PowerShell ISE)
   - Testar imediatamente após a conversão

## Comandos Úteis para Futuros Problemas

### Verificar Encoding de um Arquivo

```powershell
# No PowerShell
$bytes = [System.IO.File]::ReadAllBytes("arquivo.ps1")
if ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    Write-Host "UTF-8 com BOM"
} else {
    Write-Host "Outro encoding"
}
```

### Ver Histórico de Alterações

```powershell
# Ver mudanças entre commits
git diff HEAD~2 HEAD -- install.ps1

# Ver versão específica de um arquivo
git show 210f223:install.ps1

# Restaurar arquivo de commit específico
git checkout 210f223 -- install.ps1
```

### Testar Scripts PowerShell

```powershell
# Testar carregamento
PowerShell -NoProfile -ExecutionPolicy Bypass -Command ". .\script.ps1"

# Ver ajuda
Get-Help .\script.ps1 -Detailed

# Verificar sintaxe
$null = [System.Management.Automation.PSParser]::Tokenize(
    (Get-Content .\script.ps1 -Raw), [ref]$null
)
```

## Referências

- **Commit com correção**: 2ccbbba
- **Último commit funcional**: 210f223
- **Commits problemáticos**: 31efecc, 2ed7d37

## Próximos Passos

Se houver necessidade de alterar o encoding dos arquivos no futuro:

1. Criar um branch de teste
2. Fazer a alteração de encoding usando VS Code (salvar com encoding específico)
3. Testar **TODOS** os scripts afetados
4. Validar que caracteres especiais estão corretos
5. Apenas então fazer commit e merge

---

**Status**: ✅ Problema resolvido  
**Data da correção**: 07/11/2024  
**Responsável**: Christian Martin dos Santos
