# Script de Instalação - Chainsaw

## 📋 Visão Geral

O script `install.ps1` automatiza completamente a instalação das configurações do Word para o sistema Chainsaw, garantindo uma instalação segura e consistente para todos os usuários.

## ⚠️ IMPORTANTE: Privilégios de Administrador

**🚫 NÃO EXECUTE COMO ADMINISTRADOR**

Este script foi projetado para rodar com privilégios de **usuário normal** e:

- ✅ **Funciona perfeitamente** sem privilégios de administrador
- ✅ Opera apenas em pastas do perfil do usuário
- ✅ Não modifica arquivos do sistema
- ✅ Não requer acesso a recursos protegidos

**❌ Executar como Administrador pode causar problemas:**

- Arquivos criados com proprietário "Administrador"
- Problemas de permissões para acessar os arquivos depois
- Word pode não conseguir acessar os templates
- Operação desnecessária e insegura

## 🎯 O que o Script Faz

### 1. Validação Pré-instalação
- ✅ Verifica versão do Windows (10+)
- ✅ Verifica versão do PowerShell (5.1+)
- ✅ Confirma acesso à rede corporativa
- ✅ Testa permissões de escrita no perfil do usuário
- ✅ Valida existência dos arquivos de origem

### 2. Backup Automático
- 🔒 Renomeia pasta Templates existente com timestamp
- 🔒 Mantém histórico dos últimos 5 backups
- 🔒 Remove backups antigos automaticamente
- 🔒 Formato: `Templates_backup_YYYYMMDD_HHMMSS`

### 3. Instalação
- 📁 Copia `stamp.png` para `%USERPROFILE%\chainsaw\assets\`
- 📁 Copia Templates para `%APPDATA%\Microsoft\Templates\`
- 📁 Preserva toda estrutura de pastas e arquivos
- 📁 Verifica integridade dos arquivos copiados

### 4. Sistema de Log
- 📝 Registra todas as operações
- 📝 Salva em `%USERPROFILE%\chainsaw\logs\`
- 📝 Formato: `install_YYYYMMDD_HHMMSS.log`
- 📝 Inclui timestamps, níveis e mensagens detalhadas

### 5. Tratamento de Erros
- 🛡️ Validação completa antes de iniciar
- 🛡️ Rollback automático em caso de falha
- 🛡️ Mensagens de erro claras e acionáveis
- 🛡️ Não interrompe em avisos não críticos

## � Bypass Automático de Execução (Novo!)

O script agora possui um **mecanismo de auto-relançamento seguro** que elimina a necessidade de configurar manualmente a política de execução do PowerShell.

### Como Funciona

1. **Detecção Automática**: O script detecta se a política de execução impede sua execução
2. **Informação Clara**: Exibe informações de segurança sobre o que será feito
3. **Relançamento Seguro**: Relança-se automaticamente com `-ExecutionPolicy Bypass`
4. **Temporário**: O bypass é válido APENAS para esta execução do script
5. **Sem Alterações**: A política do sistema permanece inalterada
6. **Sem Admin**: Nenhum privilégio de administrador é necessário

### Garantias de Segurança

✅ **Isolado**: Apenas este script específico é executado com bypass  
✅ **Temporário**: O bypass expira automaticamente quando o script termina  
✅ **Transparente**: Todas as ações são informadas ao usuário  
✅ **Auditável**: Tudo é registrado no arquivo de log  
✅ **Sem Admin**: Não requer nem usa privilégios elevados  
✅ **Reversível**: A política original permanece intacta  

### Uso

Simplesmente execute o script normalmente:

```powershell
cd "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw"
.\install.ps1
```

Se necessário, o script se relançará automaticamente. Você verá:

```
🔒 Verificando política de execução...
   Política atual (CurrentUser): Restricted
⚠  Política de execução restritiva detectada.
🔄 Relançando script com bypass temporário...

ℹ  SEGURANÇA:
   • Apenas ESTE script será executado com bypass
   • A política do sistema NÃO será alterada
   • O bypass expira quando o script terminar
   • Nenhum privilégio de administrador é usado

✓ Executando com bypass temporário (seguro)
```

## �🚀 Como Usar

### Verificação de Privilégios (Obrigatória)

**PRIMEIRO: Verifique se você NÃO está executando como Administrador**

```powershell
# Execute este comando para verificar:
[bool]([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Se retornar "True", você ESTÁ como Admin (errado!)
# Se retornar "False", você está como usuário normal (correto!)
```

**Como abrir PowerShell SEM privilégios de administrador:**

1. Pressione `Win + X`
2. Selecione **"Windows PowerShell"** (NÃO selecione "Windows PowerShell (Admin)")
3. Ou simplesmente pesquise "PowerShell" no menu Iniciar e abra normalmente

### Teste de Permissões (Recomendado)

Antes de instalar, execute o script de teste de permissões:

```powershell
cd "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw"
.\test-permissions.ps1
```

Este script verifica:
- ✅ Se você NÃO está executando como administrador
- ✅ Permissões de escrita em `%USERPROFILE%`
- ✅ Permissões de escrita em `%APPDATA%`
- ✅ Capacidade de criar, renomear e copiar arquivos/pastas

### Teste Rápido de Instalação (Opcional)

Para verificar pré-requisitos sem modificar nada:

```powershell
cd "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw"
.\test-install.ps1
```

### Instalação Interativa (Padrão)

**Método Recomendado - Usando o Launcher Seguro:**

```cmd
cd "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw"
install.cmd
```

**Alternativa - Execução Direta do PowerShell:**

```powershell
cd "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw"
.\install.ps1
```

O script irá:

1. **Verificar e ajustar automaticamente a política de execução** (bypass temporário seguro)
2. Verificar pré-requisitos
3. Mostrar o que será feito
4. Pedir confirmação
5. Executar a instalação
6. Exibir resultado detalhado

🔒 **Segurança do Bypass Automático:**

- ✅ Apenas ESTE script é executado com bypass
- ✅ A política do sistema NÃO é alterada permanentemente
- ✅ O bypass expira automaticamente quando o script termina
- ✅ Nenhum privilégio de administrador é necessário ou usado
- ✅ Totalmente transparente e seguro
- ✅ O launcher `.cmd` funciona em QUALQUER política de execução

### Instalação Automática

Para instalação sem interação (útil para scripts de deploy):

```cmd
install.cmd -Force
```

Ou diretamente:

```powershell
.\install.ps1 -Force
```

### Instalação Sem Backup

⚠️ **Não recomendado** - Instala sem criar backup:

```cmd
install.cmd -NoBackup
```

### Instalação com Caminho Customizado

```cmd
install.cmd -SourcePath "\\outro-servidor\caminho\chainsaw"
```

## 📊 Exemplo de Execução

```
╔════════════════════════════════════════════════════════════════╗
║          CHAINSAW - Instalação de Configurações do Word       ║
╚════════════════════════════════════════════════════════════════╝

📝 Arquivo de log: C:\Users\csantos\chainsaw\logs\install_20251105_143022.log

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ETAPA 1: Verificação de Pré-requisitos
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ℹ Verificando pré-requisitos...
✓ Sistema operacional: Windows 10.0 ✓
✓ PowerShell versão: 5.1.19041.4894 ✓
ℹ Verificando acesso ao caminho de rede: \\strqnapmain\Dir. Legislativa\_Christian261\chainsaw
✓ Acesso ao caminho de rede confirmado ✓
✓ Permissões de escrita no perfil do usuário confirmadas ✓

[... mais output ...]

╔════════════════════════════════════════════════════════════════╗
║              INSTALAÇÃO CONCLUÍDA COM SUCESSO!                 ║
╚════════════════════════════════════════════════════════════════╝

📊 Resumo da Instalação:
   • Operações bem-sucedidas: 5
   • Avisos: 0
   • Erros: 0
   • Tempo decorrido: 00:12
```

## 🔍 Estrutura do Log

```
================================================================================
CHAINSAW - Log de Instalação
================================================================================
Data/Hora Início: 05/11/2025 14:30:22
Usuário: csantos
Computador: DESKTOP-ABC123
Sistema: Microsoft Windows NT 10.0.19045.0
PowerShell: 5.1.19041.4894
Caminho de Origem: \\strqnapmain\Dir. Legislativa\_Christian261\chainsaw
================================================================================

[2025-11-05 14:30:22] [INFO] === INÍCIO DA INSTALAÇÃO ===
[2025-11-05 14:30:22] [INFO] Verificando pré-requisitos...
[2025-11-05 14:30:22] [SUCCESS] Sistema operacional: Windows 10.0 ✓
[2025-11-05 14:30:23] [SUCCESS] PowerShell versão: 5.1.19041.4894 ✓
[2025-11-05 14:30:23] [INFO] Verificando acesso ao caminho de rede: ...
[2025-11-05 14:30:24] [SUCCESS] Acesso ao caminho de rede confirmado ✓
...
```

## 🛠️ Solução de Problemas

### Erro: Script não pode ser executado

**Problema:** "O arquivo install.ps1 não pode ser carregado porque a execução de scripts está desabilitada neste sistema."

**Solução Automática (Recomendada):**

O script `install.ps1` **detecta automaticamente** este problema e se relança com bypass temporário. Simplesmente execute:

```powershell
.\install.ps1
```

O script irá:
1. Detectar a política restritiva
2. Mostrar informações de segurança
3. Relançar-se automaticamente com bypass temporário
4. Executar a instalação normalmente
5. Retornar à política original automaticamente

**Solução Manual (Alternativa):**

Se preferir configurar manualmente a política de execução de forma permanente:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

⚠️ **Nota:** A solução automática é mais segura, pois não altera permanentemente as configurações do sistema.

### Erro: Caminho de rede não acessível

**Problema:** "Não foi possível acessar o caminho de rede"

**Possíveis causas:**
1. Não está conectado à VPN/rede corporativa
2. Credenciais de rede expiradas
3. Caminho incorreto ou servidor offline

**Solução:**
1. Conecte-se à VPN/rede corporativa
2. Teste o acesso manualmente: `explorer "\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw"`
3. Verifique suas credenciais de rede

### Erro: Permissões insuficientes

**Problema:** "Sem permissões de escrita no perfil do usuário"

**Solução:**
1. **NÃO** execute como Administrador
2. Execute como seu usuário normal
3. Verifique se não há restrições de política de grupo

### Word em Execução

**Problema:** Avisos sobre Word em execução

**Solução:**
1. Feche completamente o Microsoft Word
2. Feche todos os documentos do Office
3. Verifique no Gerenciador de Tarefas se `WINWORD.EXE` está em execução
4. Se persistir, reinicie o computador

### Erro na Cópia de Arquivos

**Problema:** "Erro ao copiar pasta Templates"

**Possíveis causas:**
1. Arquivos bloqueados pelo Word
2. Antivírus bloqueando acesso
3. Disco cheio

**Solução:**
1. Feche o Word completamente
2. Adicione exceção no antivírus para a pasta Templates
3. Verifique espaço em disco: `Get-PSDrive C`

## 🔐 Segurança

### O que o script NÃO faz

- ❌ Não requer privilégios de administrador
- ❌ Não modifica arquivos do sistema
- ❌ Não altera registro do Windows
- ❌ Não instala software adicional
- ❌ Não faz comunicação externa
- ❌ Não coleta dados do usuário

### O que o script faz para segurança

- ✅ Valida todos os inputs
- ✅ Cria backup antes de modificar
- ✅ Registra todas as operações em log
- ✅ Reverte mudanças em caso de erro
- ✅ Verifica integridade dos arquivos
- ✅ Opera apenas no perfil do usuário

## 📁 Estrutura de Arquivos Criada

Após a instalação, a seguinte estrutura será criada:

```
%USERPROFILE%\
├─ chainsaw\
│  ├─ assets\
│  │  └─ stamp.png              # Imagem do cabeçalho
│  └─ logs\
│     └─ install_*.log          # Logs de instalação
│
%APPDATA%\Microsoft\
├─ Templates\                    # Configurações do Word
│  ├─ LiveContent\
│  │  └─ 16\
│  │     └─ Managed\
│  │        ├─ Document Themes\
│  │        ├─ SmartArt Graphics\
│  │        ├─ Word Document Bibliography Styles\
│  │        └─ Word Document Building Blocks\
│  └─ ...
│
└─ Templates_backup_YYYYMMDD_HHMMSS\  # Backup da instalação anterior
   └─ [conteúdo anterior]
```

## 🔄 Atualizações

Para atualizar uma instalação existente:

1. Execute `.\install.ps1` novamente
2. O script criará um novo backup automático
3. As configurações antigas serão preservadas no backup
4. As novas configurações serão instaladas

## 📞 Suporte

Se encontrar problemas não listados aqui:

1. Consulte o arquivo de log: `%USERPROFILE%\chainsaw\logs\install_*.log`
2. Execute `.\test-install.ps1` para diagnóstico
3. Verifique o README.md principal para documentação completa
4. Entre em contato com Christian Martin (chrmsantos@protonmail.com)

## 📜 Licença

GNU General Public License v3.0 (GPLv3)

---

**Versão:** 1.0.0  
**Última Atualização:** 05/11/2025  
**Autor:** Christian Martin dos Santos
