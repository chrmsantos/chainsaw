# 🔐 Garantia de Execução sem Privilégios de Administrador

## ✅ Confirmação

O script de instalação do Chainsaw (`install.ps1`) **FUNCIONA COMPLETAMENTE** sem privilégios de administrador e foi projetado especificamente para isso.

## 📍 Áreas de Operação

O script opera EXCLUSIVAMENTE nas seguintes áreas do perfil do usuário:

### 1. `%USERPROFILE%\chainsaw\`
- **Caminho completo:** `C:\Users\[seu_usuario]\chainsaw\`
- **Permissões:** Usuário normal tem controle total
- **Operações:**
  - Criar pasta `chainsaw`
  - Criar subpasta `assets`
  - Criar subpasta `logs`
  - Copiar arquivo `stamp.png`
  - Criar arquivos de log

### 2. `%APPDATA%\Microsoft\Templates\`
- **Caminho completo:** `C:\Users\[seu_usuario]\AppData\Roaming\Microsoft\Templates\`
- **Permissões:** Usuário normal tem controle total
- **Operações:**
  - Renomear pasta existente (backup)
  - Criar nova pasta Templates
  - Copiar toda estrutura de arquivos e pastas
  - Manter backups antigos

## 🚫 O que o Script NÃO Faz

O script foi projetado para **NÃO** realizar nenhuma das seguintes operações que requerem privilégios elevados:

- ❌ Não modifica `C:\Windows\`
- ❌ Não modifica `C:\Program Files\`
- ❌ Não modifica `C:\Program Files (x86)\`
- ❌ Não modifica o Registro do Windows
- ❌ Não cria serviços do Windows
- ❌ Não instala drivers
- ❌ Não modifica políticas de grupo
- ❌ Não modifica configurações de firewall
- ❌ Não acessa pastas de outros usuários
- ❌ Não modifica permissões de arquivos
- ❌ Não executa comandos do sistema

## 🛡️ Proteções Implementadas

### 1. Verificação Ativa
O script verifica se está sendo executado como administrador e:
- Exibe aviso visual destacado
- Explica os problemas que podem ocorrer
- Pede confirmação explícita para continuar
- Recomenda fortemente executar como usuário normal

### 2. Teste de Permissões
Script `test-permissions.ps1` verifica:
- ✅ Modo de execução (deve ser usuário normal)
- ✅ Permissões de escrita em `%USERPROFILE%`
- ✅ Permissões de escrita em `%APPDATA%`
- ✅ Criação de diretórios
- ✅ Renomeação de pastas
- ✅ Cópia de arquivos
- ✅ Cópia recursiva de diretórios

## ⚙️ Operações Realizadas e Permissões Necessárias

| Operação | Local | Permissão Necessária | Admin? |
|----------|-------|---------------------|--------|
| Criar pasta `chainsaw` | `%USERPROFILE%` | Escrita no perfil | ❌ NÃO |
| Copiar `stamp.png` | `%USERPROFILE%\chainsaw\assets` | Escrita no perfil | ❌ NÃO |
| Criar logs | `%USERPROFILE%\chainsaw\logs` | Escrita no perfil | ❌ NÃO |
| Renomear Templates | `%APPDATA%\Microsoft` | Escrita em AppData | ❌ NÃO |
| Copiar Templates | `%APPDATA%\Microsoft` | Escrita em AppData | ❌ NÃO |
| Ler da rede | `\\servidor\caminho` | Acesso à rede | ❌ NÃO |

## ❌ Por Que NÃO Executar como Administrador?

### Problema 1: Propriedade de Arquivos
Se executado como administrador:
- Arquivos são criados com proprietário "Administrador"
- Seu usuário normal pode ter problemas para acessá-los
- Word pode não conseguir ler os templates

### Problema 2: Perfil Incorreto
Se executado como administrador:
- `%USERPROFILE%` pode apontar para `C:\Users\Administrador`
- Arquivos seriam instalados no perfil errado
- Seu usuário não teria acesso

### Problema 3: Segurança
- Executar scripts com privilégios elevados é uma má prática de segurança
- Aumenta superfície de ataque
- Não há necessidade real

## ✅ Como Garantir Execução Correta

### Passo 1: Abrir PowerShell Corretamente

**MÉTODO 1 - Recomendado:**
1. Pressione `Win + X`
2. Clique em **"Windows PowerShell"**
3. NÃO clique em "Windows PowerShell (Admin)"

**MÉTODO 2:**
1. Pressione `Win + R`
2. Digite: `powershell`
3. Pressione Enter

**MÉTODO 3:**
1. Abra o Menu Iniciar
2. Digite: `PowerShell`
3. Clique normalmente (não clique com botão direito)

### Passo 2: Verificar Status

Execute este comando para verificar:

```powershell
[bool]([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
```

**Resultado esperado:** `False`

Se retornar `True`, você está como administrador. Feche e abra novamente sem privilégios.

### Passo 3: Testar Permissões

```powershell
cd "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw"
.\test-permissions.ps1
```

Todos os testes devem passar ✓

### Passo 4: Executar Instalação

```powershell
.\install.ps1
```

O script verificará automaticamente e avisará se detectar privilégios de administrador.

## 📊 Validação Técnica

### Comandos Utilizados

Todos os comandos do PowerShell utilizados no script funcionam sem privilégios de administrador:

- ✅ `New-Item` - Criar pastas/arquivos no perfil do usuário
- ✅ `Copy-Item` - Copiar arquivos/pastas
- ✅ `Rename-Item` - Renomear pastas
- ✅ `Remove-Item` - Remover arquivos/pastas do usuário
- ✅ `Test-Path` - Verificar existência de arquivos
- ✅ `Get-ChildItem` - Listar arquivos/pastas
- ✅ `Get-FileHash` - Calcular hash de arquivos
- ✅ `Write-Host` - Escrever na tela
- ✅ `Add-Content` - Adicionar conteúdo a arquivos
- ✅ `Join-Path` - Construir caminhos
- ✅ `Split-Path` - Dividir caminhos
- ✅ `Get-Date` - Obter data/hora

### Variáveis de Ambiente

Todas as variáveis de ambiente utilizadas são acessíveis ao usuário normal:

- ✅ `$env:USERPROFILE` - Perfil do usuário atual
- ✅ `$env:APPDATA` - AppData\Roaming do usuário
- ✅ `$env:USERNAME` - Nome do usuário
- ✅ `$env:COMPUTERNAME` - Nome do computador
- ✅ `$env:TEMP` - Pasta temporária do usuário

### .NET Framework Classes

Todas as classes .NET utilizadas são acessíveis:

- ✅ `[System.IO.File]` - Operações com arquivos
- ✅ `[System.IO.Directory]` - Operações com diretórios
- ✅ `[Environment]` - Informações do ambiente
- ✅ `[Security.Principal.WindowsPrincipal]` - Verificação de identidade

## 🧪 Testes Realizados

Todos os seguintes testes foram implementados em `test-permissions.ps1`:

1. ✅ Verificação de modo de execução (não admin)
2. ✅ Escrita em `%USERPROFILE%`
3. ✅ Criação de diretórios em `%USERPROFILE%`
4. ✅ Escrita em `%APPDATA%`
5. ✅ Renomeação de pastas em `%APPDATA%`
6. ✅ Cópia de arquivos individuais
7. ✅ Cópia recursiva de diretórios com estrutura
8. ✅ Acesso a informações do sistema

## 📝 Conclusão

O script de instalação do Chainsaw:

✅ **GARANTE** execução sem privilégios de administrador
✅ **OPERA** exclusivamente no perfil do usuário
✅ **VERIFICA** ativamente se está sendo executado como admin
✅ **AVISA** claramente sobre problemas de execução elevada
✅ **TESTA** todas as permissões necessárias
✅ **DOCUMENTA** completamente todos os requisitos
✅ **IMPLEMENTA** todas as melhores práticas de segurança

---

**Status:** ✅ CERTIFICADO PARA EXECUÇÃO SEM PRIVILÉGIOS DE ADMINISTRADOR

**Versão:** 1.0.0  
**Data:** 05/11/2025  
**Autor:** Christian Martin dos Santos
