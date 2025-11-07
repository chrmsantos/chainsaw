# [SEC] Garantia de Execução sem Privilégios de Administrador

## [OK] Confirmação

O script de instalação do Chainsaw (`install.ps1`) **FUNCIONA COMPLETAMENTE** sem privilégios de administrador e foi projetado especificamente para isso.

## 📍 Áreas de Operação

O script opera EXCLUSIVAMENTE nas seguintes áreas do perfil do usuário:

### 1. `%USERPROFILE%\chainsaw\`
- **Caminho completo:** ` # [SEC] Garantia de Execução sem Privilégios de Administrador

## [OK] Confirmação

O script de instalação do Chainsaw (`install.ps1`) **FUNCIONA COMPLETAMENTE** sem privilégios de administrador e foi projetado especificamente para isso.

## 📍 Áreas de Operação

O script opera EXCLUSIVAMENTE nas seguintes áreas do perfil do usuário:

### 1. `%USERPROFILE%\chainsaw\`
- **Caminho completo:** `C:\Users\[seu_usuario]\chainsaw\`
- **Permissões:** Usuário normal tem controle total
- **Operações:**
  - Criar Pasta `chainsaw`
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

## [NO] O que o Script NÃO Faz

O script foi projetado para **NÃO** realizar nenhuma das seguintes operações que requerem privilégios elevados:

- [X] Não modifica `C:\Windows\`
- [X] Não modifica `C:\Program Files\`
- [X] Não modifica `C:\Program Files (x86)\`
- [X] Não modifica o Registro do Windows
- [X] Não cria serviços do Windows
- [X] Não instala drivers
- [X] Não modifica políticas de grupo
- [X] Não modifica configurações de firewall
- [X] Não acessa pastas de outros usuários
- [X] Não modifica permissões de arquivos
- [X] Não executa comandos do sistema

## [SEC] Proteções Implementadas

### 1. Verificação Ativa
O script verifica se está sendo executado como administrador e:
- Exibe aviso visual destacado
- Explica os problemas que podem ocorrer
- Pede confirmação explícita para continuar
- Recomenda fortemente executar como usuário normal

### 2. Teste de Permissões
Script `test-permissions.ps1` verifica:
- [OK] Modo de execução (deve ser usuário normal)
- [OK] Permissões de escrita em `%USERPROFILE%`
- [OK] Permissões de escrita em `%APPDATA%`
- [OK] Criação de diretórios
- [OK] Renomeação de pastas
- [OK] Cópia de arquivos
- [OK] Cópia recursiva de diretórios

## ⚙️ Operações Realizadas e Permissões Necessárias

| Operação | Local | Permissão Necessária | Admin? |
|----------|-------|---------------------|--------|
| Criar Pasta `chainsaw` | `%USERPROFILE%` | Escrita no perfil | [X] NÃO |
| Copiar `stamp.png` | `%USERPROFILE%\chainsaw\assets` | Escrita no perfil | [X] NÃO |
| Criar logs | `%USERPROFILE%\chainsaw\logs` | Escrita no perfil | [X] NÃO |
| Renomear Templates | `%APPDATA%\Microsoft` | Escrita em AppData | [X] NÃO |
| Copiar Templates | `%APPDATA%\Microsoft` | Escrita em AppData | [X] NÃO |
| Ler da rede | `\\servidor\caminho` | Acesso à rede | [X] NÃO |

## [X] Por Que NÃO Executar como Administrador?

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

## [OK] Como Garantir Execução Correta

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

Todos os testes devem passar [OK]

### Passo 4: Executar Instalação

```powershell
.\install.ps1
```

O script verificará automaticamente e avisará se detectar privilégios de administrador.

## [CHART] Validação Técnica

### Comandos Utilizados

Todos os comandos do PowerShell utilizados no script funcionam sem privilégios de administrador:

- [OK] `New-Item` - Criar pastas/arquivos no perfil do usuário
- [OK] `Copy-Item` - Copiar arquivos/pastas
- [OK] `Rename-Item` - Renomear pastas
- [OK] `Remove-Item` - Remover arquivos/pastas do usuário
- [OK] `Test-Path` - Verificar existência de arquivos
- [OK] `Get-ChildItem` - Listar arquivos/pastas
- [OK] `Get-FileHash` - Calcular hash de arquivos
- [OK] `Write-Host` - Escrever na tela
- [OK] `Add-Content` - Adicionar conteúdo a arquivos
- [OK] `Join-Path` - Construir caminhos
- [OK] `Split-Path` - Dividir caminhos
- [OK] `Get-Date` - Obter data/hora

### Variáveis de Ambiente

Todas as variáveis de ambiente utilizadas são acessíveis ao usuário normal:

- [OK] `$env:USERPROFILE` - Perfil do usuário atual
- [OK] `$env:APPDATA` - AppData\Roaming do usuário
- [OK] `$env:USERNAME` - Nome do usuário
- [OK] `$env:COMPUTERNAME` - Nome do computador
- [OK] `$env:TEMP` - Pasta temporária do usuário

### .NET Framework Classes

Todas as classes .NET utilizadas são acessíveis:

- [OK] `[System.IO.File]` - Operações com arquivos
- [OK] `[System.IO.Directory]` - Operações com diretórios
- [OK] `[Environment]` - Informações do ambiente
- [OK] `[Security.Principal.WindowsPrincipal]` - Verificação de identidade

## 🧪 Testes Realizados

Todos os seguintes testes foram implementados em `test-permissions.ps1`:

1. [OK] Verificação de modo de execução (não admin)
2. [OK] Escrita em `%USERPROFILE%`
3. [OK] Criação de diretórios em `%USERPROFILE%`
4. [OK] Escrita em `%APPDATA%`
5. [OK] Renomeação de pastas em `%APPDATA%`
6. [OK] Cópia de arquivos individuais
7. [OK] Cópia recursiva de diretórios com estrutura
8. [OK] Acesso a informações do sistema

## [LOG] Conclusão

O script de instalação do Chainsaw:

[OK] **GARANTE** execução sem privilégios de administrador
[OK] **OPERA** exclusivamente no perfil do usuário
[OK] **VERIFICA** ativamente se está sendo executado como admin
[OK] **AVISA** claramente sobre problemas de execução elevada
[OK] **TESTA** todas as permissões necessárias
[OK] **DOCUMENTA** completamente todos os requisitos
[OK] **IMPLEMENTA** todas as melhores práticas de segurança

---

**Status:** [OK] CERTIFICADO PARA EXECUÇÃO SEM PRIVILÉGIOS DE ADMINISTRADOR

**Versão:** 1.0.0  
**Data:** 05/11/2025  
**Autor:** Christian Martin dos Santos
.Value -replace 'CHAINSAW', 'chainsaw' \`
- **Permissões:** Usuário normal tem controle total
- **Operações:**
  - Criar Pasta `chainsaw`
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

## [NO] O que o Script NÃO Faz

O script foi projetado para **NÃO** realizar nenhuma das seguintes operações que requerem privilégios elevados:

- [X] Não modifica `C:\Windows\`
- [X] Não modifica `C:\Program Files\`
- [X] Não modifica `C:\Program Files (x86)\`
- [X] Não modifica o Registro do Windows
- [X] Não cria serviços do Windows
- [X] Não instala drivers
- [X] Não modifica políticas de grupo
- [X] Não modifica configurações de firewall
- [X] Não acessa pastas de outros usuários
- [X] Não modifica permissões de arquivos
- [X] Não executa comandos do sistema

## [SEC] Proteções Implementadas

### 1. Verificação Ativa
O script verifica se está sendo executado como administrador e:
- Exibe aviso visual destacado
- Explica os problemas que podem ocorrer
- Pede confirmação explícita para continuar
- Recomenda fortemente executar como usuário normal

### 2. Teste de Permissões
Script `test-permissions.ps1` verifica:
- [OK] Modo de execução (deve ser usuário normal)
- [OK] Permissões de escrita em `%USERPROFILE%`
- [OK] Permissões de escrita em `%APPDATA%`
- [OK] Criação de diretórios
- [OK] Renomeação de pastas
- [OK] Cópia de arquivos
- [OK] Cópia recursiva de diretórios

## ⚙️ Operações Realizadas e Permissões Necessárias

| Operação | Local | Permissão Necessária | Admin? |
|----------|-------|---------------------|--------|
| Criar Pasta `chainsaw` | `%USERPROFILE%` | Escrita no perfil | [X] NÃO |
| Copiar `stamp.png` | `%USERPROFILE%\chainsaw\assets` | Escrita no perfil | [X] NÃO |
| Criar logs | `%USERPROFILE%\chainsaw\logs` | Escrita no perfil | [X] NÃO |
| Renomear Templates | `%APPDATA%\Microsoft` | Escrita em AppData | [X] NÃO |
| Copiar Templates | `%APPDATA%\Microsoft` | Escrita em AppData | [X] NÃO |
| Ler da rede | `\\servidor\caminho` | Acesso à rede | [X] NÃO |

## [X] Por Que NÃO Executar como Administrador?

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

## [OK] Como Garantir Execução Correta

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

Todos os testes devem passar [OK]

### Passo 4: Executar Instalação

```powershell
.\install.ps1
```

O script verificará automaticamente e avisará se detectar privilégios de administrador.

## [CHART] Validação Técnica

### Comandos Utilizados

Todos os comandos do PowerShell utilizados no script funcionam sem privilégios de administrador:

- [OK] `New-Item` - Criar pastas/arquivos no perfil do usuário
- [OK] `Copy-Item` - Copiar arquivos/pastas
- [OK] `Rename-Item` - Renomear pastas
- [OK] `Remove-Item` - Remover arquivos/pastas do usuário
- [OK] `Test-Path` - Verificar existência de arquivos
- [OK] `Get-ChildItem` - Listar arquivos/pastas
- [OK] `Get-FileHash` - Calcular hash de arquivos
- [OK] `Write-Host` - Escrever na tela
- [OK] `Add-Content` - Adicionar conteúdo a arquivos
- [OK] `Join-Path` - Construir caminhos
- [OK] `Split-Path` - Dividir caminhos
- [OK] `Get-Date` - Obter data/hora

### Variáveis de Ambiente

Todas as variáveis de ambiente utilizadas são acessíveis ao usuário normal:

- [OK] `$env:USERPROFILE` - Perfil do usuário atual
- [OK] `$env:APPDATA` - AppData\Roaming do usuário
- [OK] `$env:USERNAME` - Nome do usuário
- [OK] `$env:COMPUTERNAME` - Nome do computador
- [OK] `$env:TEMP` - Pasta temporária do usuário

### .NET Framework Classes

Todas as classes .NET utilizadas são acessíveis:

- [OK] `[System.IO.File]` - Operações com arquivos
- [OK] `[System.IO.Directory]` - Operações com diretórios
- [OK] `[Environment]` - Informações do ambiente
- [OK] `[Security.Principal.WindowsPrincipal]` - Verificação de identidade

## 🧪 Testes Realizados

Todos os seguintes testes foram implementados em `test-permissions.ps1`:

1. [OK] Verificação de modo de execução (não admin)
2. [OK] Escrita em `%USERPROFILE%`
3. [OK] Criação de diretórios em `%USERPROFILE%`
4. [OK] Escrita em `%APPDATA%`
5. [OK] Renomeação de pastas em `%APPDATA%`
6. [OK] Cópia de arquivos individuais
7. [OK] Cópia recursiva de diretórios com estrutura
8. [OK] Acesso a informações do sistema

## [LOG] Conclusão

O script de instalação do Chainsaw:

[OK] **GARANTE** execução sem privilégios de administrador
[OK] **OPERA** exclusivamente no perfil do usuário
[OK] **VERIFICA** ativamente se está sendo executado como admin
[OK] **AVISA** claramente sobre problemas de execução elevada
[OK] **TESTA** todas as permissões necessárias
[OK] **DOCUMENTA** completamente todos os requisitos
[OK] **IMPLEMENTA** todas as melhores práticas de segurança

---

**Status:** [OK] CERTIFICADO PARA EXECUÇÃO SEM PRIVILÉGIOS DE ADMINISTRADOR

**Versão:** 1.0.0  
**Data:** 05/11/2025  
**Autor:** Christian Martin dos Santos
