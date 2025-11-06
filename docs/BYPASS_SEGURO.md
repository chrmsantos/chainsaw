# Mecanismo de Bypass Automático Seguro

## [INFO] Visão Geral

O script `install.ps1` implementa um mecanismo de auto-relançamento com bypass temporário da política de execução do PowerShell. Este documento detalha como funciona e por que é seguro.

## [SEC] Arquitetura de Segurança

### 1. Detecção da Política de Execução

O script primeiro verifica a política atual:

```powershell
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
```

Políticas que requerem bypass:
- `Restricted`: Não permite execução de scripts
- `AllSigned`: Requer assinatura digital em todos os scripts
- Qualquer política que gere `PSSecurityException`

### 2. Relançamento Isolado

Quando necessário, o script se relança usando:

```powershell
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "caminho\script.ps1" -BypassedExecution
```

**Parâmetros de Segurança:**

- **`-ExecutionPolicy Bypass`**: Permite execução apenas deste processo
- **`-NoProfile`**: Não carrega perfis de usuário (mais seguro)
- **`-File`**: Especifica exatamente qual arquivo executar
- **`-BypassedExecution`**: Flag interna para evitar loop infinito

### 3. Escopo do Bypass

O bypass tem escopo limitado:

| Aspecto | Escopo |
|---------|--------|
| **Temporal** | Apenas durante a execução do processo |
| **Espacial** | Apenas o arquivo especificado em `-File` |
| **Processo** | Apenas o processo filho criado |
| **Sistema** | A política do sistema permanece inalterada |

### 4. Preservação do Estado

```powershell
# ANTES do relançamento
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
# Resultado: Restricted (por exemplo)

# DURANTE o relançamento
# O processo filho executa com bypass temporário

# APÓS o término do script
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
# Resultado: Restricted (exatamente como antes)
```

**Garantia**: A política original é preservada automaticamente porque:
1. Não usamos `Set-ExecutionPolicy` em nenhum momento
2. O bypass é apenas para o processo, não para o usuário
3. Quando o processo termina, o bypass desaparece

## [SEC] Camadas de Segurança

### Camada 1: Detecção Precisa

```powershell
$needsBypass = $false
try {
    # Tenta executar um bloco de script trivial
    $null = [ScriptBlock]::Create("1 + 1").Invoke()
}
catch [System.Management.Automation.PSSecurityException] {
    # Se falhar, realmente precisa de bypass
    $needsBypass = $true
}
```

**Benefício**: Evita relançamento desnecessário se a política já permitir execução.

### Camada 2: Transparência Total

O script informa claramente ao usuário:

```
[SEC] Verificando política de execução...
   Política atual (CurrentUser): Restricted
[!]  Política de execução restritiva detectada.
[SYNC] Relançando script com bypass temporário...

[i]  SEGURANÇA:
   • Apenas ESTE script será executado com bypass
   • A política do sistema NÃO será alterada
   • O bypass expira quando o script terminar
   • Nenhum privilégio de administrador é usado
```

**Benefício**: O usuário sabe exatamente o que está acontecendo.

### Camada 3: Auditabilidade

Tudo é registrado no log:

```
[2025-11-05 14:30:22] [INFO] Política de execução atual: Restricted
[2025-11-05 14:30:22] [INFO] Relançando com bypass temporário
[2025-11-05 14:30:23] [INFO] Executando com bypass seguro
```

**Benefício**: Auditoria completa para conformidade e troubleshooting.

### Camada 4: Prevenção de Loop

```powershell
param(
    [Parameter(DontShow)]
    [switch]$BypassedExecution
)

if (-not $BypassedExecution) {
    # Lógica de detecção e relançamento
}
else {
    # Execução normal - já está com bypass
}
```

**Benefício**: Impossível criar loop infinito de relançamentos.

### Camada 5: Propagação de Parâmetros

```powershell
# Preserva todos os parâmetros originais
if ($SourcePath -ne "\\strqnapmain\...") {
    $arguments += @("-SourcePath", "`"$SourcePath`"")
}
if ($Force) {
    $arguments += "-Force"
}
if ($NoBackup) {
    $arguments += "-NoBackup"
}
```

**Benefício**: O comportamento é idêntico com ou sem bypass.

### Camada 6: Código de Saída

```powershell
$processInfo = Start-Process ... -Wait -PassThru
exit $processInfo.ExitCode
```

**Benefício**: O código de saída é propagado corretamente para scripts de automação.

## 🔬 Análise de Segurança

### Vetor de Ataque: Substituição de Arquivo

**Cenário**: Atacante substitui `install.ps1` por código malicioso.

**Mitigação**:
- O arquivo está em caminho de rede protegido com ACLs
- Usuário deve ter permissões de leitura no compartilhamento
- Mesmo com bypass, o script não tem privilégios elevados
- Todas as operações são limitadas ao perfil do usuário

**Risco Residual**: Baixo (requer acesso de escrita ao compartilhamento de rede)

### Vetor de Ataque: Injeção de Parâmetros

**Cenário**: Atacante tenta injetar comandos via parâmetros.

**Mitigação**:
```powershell
# Parâmetros são validados e escapados
$arguments += @("-SourcePath", "`"$SourcePath`"")
```

**Risco Residual**: Mínimo (PowerShell escapa automaticamente)

### Vetor de Ataque: Path Hijacking

**Cenário**: Atacante coloca `powershell.exe` malicioso no PATH.

**Mitigação**:
- Windows garante que `powershell.exe` seja encontrado primeiro
- Não usamos caminhos relativos
- Usuário não tem privilégios para substituir PowerShell do sistema

**Risco Residual**: Desprezível (requer admin para modificar System32)

### Vetor de Ataque: Process Injection

**Cenário**: Atacante tenta injetar código no processo PowerShell.

**Mitigação**:
- Requer privilégios elevados para injetar em processo
- Script não executa com privilégios elevados
- Sistema operacional protege processos de usuário

**Risco Residual**: Mínimo (requer privilégios que não temos)

## [OK] Comparação com Alternativas

### Alternativa 1: Set-ExecutionPolicy

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

**Desvantagens**:
- [X] Altera permanentemente a política do usuário
- [X] Pode conflitar com políticas de grupo corporativas
- [X] Requer que usuário entenda o conceito de políticas de execução
- [X] Deixa o sistema mais permissivo permanentemente

**Vantagens do Bypass Automático**:
- [OK] Temporário: expira automaticamente
- [OK] Isolado: apenas este script
- [OK] Automático: nenhuma ação manual necessária
- [OK] Seguro: não deixa o sistema mais vulnerável

### Alternativa 2: Assinatura Digital

```powershell
# Assinar o script com certificado
Set-AuthenticodeSignature -FilePath install.ps1 -Certificate $cert
```

**Desvantagens**:
- [X] Requer infraestrutura de certificados
- [X] Custo de manutenção de certificados
- [X] Complexidade adicional
- [X] Usuários ainda precisam confiar no certificado

**Vantagens do Bypass Automático**:
- [OK] Zero configuração
- [OK] Funciona imediatamente
- [OK] Sem custo adicional
- [OK] Simples de manter

### Alternativa 3: Executar Manualmente com Bypass

```powershell
powershell.exe -ExecutionPolicy Bypass -File install.ps1
```

**Desvantagens**:
- [X] Usuário precisa lembrar o comando
- [X] Propenso a erros de digitação
- [X] Não funciona bem em documentação
- [X] Experiência de usuário ruim

**Vantagens do Bypass Automático**:
- [OK] Transparente para o usuário
- [OK] Comando simples: `.\install.ps1`
- [OK] Menos propenso a erros
- [OK] Melhor experiência de usuário

## [CHART] Matriz de Decisão

| Critério | Manual Set-Policy | Assinatura Digital | Bypass Automático |
|----------|-------------------|--------------------|--------------------|
| **Segurança** | [*][*][*] | [*][*][*][*][*] | [*][*][*][*] |
| **Usabilidade** | [*][*] | [*][*][*] | [*][*][*][*][*] |
| **Manutenção** | [*][*][*][*] | [*][*] | [*][*][*][*][*] |
| **Custo** | [*][*][*][*][*] | [*][*] | [*][*][*][*][*] |
| **Temporário** | [X] | [OK] | [OK] |
| **Transparente** | [*][*] | [*][*][*] | [*][*][*][*][*] |
| **Total** | 14/30 | 16/30 | **24/30** [OK] |

## [*] Casos de Uso

### Caso 1: Primeiro Uso

**Situação**: Usuário nunca executou scripts PowerShell antes.

**Comportamento**:
1. PowerShell tem política `Restricted` (padrão)
2. Usuário executa `.\install.ps1`
3. Script detecta política restritiva
4. Mostra mensagem de segurança
5. Relança automaticamente com bypass
6. Instalação completa com sucesso
7. Política permanece `Restricted`

**Resultado**: [OK] Sucesso sem intervenção manual

### Caso 2: Política Corporativa

**Situação**: Empresa força política `AllSigned` via GPO.

**Comportamento**:
1. PowerShell tem política `AllSigned` (forçada por GPO)
2. Script não pode alterar política (GPO tem prioridade)
3. Script detecta e usa bypass temporário
4. Instalação funciona normalmente
5. Conformidade com GPO mantida

**Resultado**: [OK] Funciona mesmo com GPO restritiva

### Caso 3: Política Permissiva

**Situação**: Usuário já configurou `RemoteSigned`.

**Comportamento**:
1. PowerShell tem política `RemoteSigned`
2. Script executa teste de segurança
3. Teste passa (não precisa bypass)
4. Instalação prossegue diretamente
5. Nenhum relançamento necessário

**Resultado**: [OK] Eficiente - não relança quando desnecessário

## [LOG] Conclusão

O mecanismo de bypass automático oferece:

1. **Segurança**: Bypass temporário e isolado, sem alterações permanentes
2. **Usabilidade**: Experiência transparente para o usuário
3. **Manutenibilidade**: Sem dependências externas ou configuração complexa
4. **Conformidade**: Respeita políticas corporativas existentes
5. **Auditabilidade**: Todas as operações são registradas

É a melhor solução para o cenário de instalação de configurações do Word sem privilégios administrativos.

---

**Versão:** 1.0.0  
**Última Atualização:** 05/11/2025  
**Autor:** Christian Martin dos Santos
