# Script de Instalação - CHAINSAW

## [INFO] Visão Geral

O script `install.ps1` automatiza completamente a instalação das configurações do Word para o sistema CHAINSAW, garantindo uma instalação segura e consistente para todos os usuários.

## [!] IMPORTANTE: Privilégios de Administrador

**[NO] NÃO EXECUTE COMO ADMINISTRADOR**

Este script foi projetado para rodar com privilégios de **usuário normal** e:

- [OK] **Funciona perfeitamente** sem privilégios de administrador
- [OK] Opera apenas em pastas do perfil do usuário
- [OK] Não modifica arquivos do sistema
- [OK] Não requer acesso a recursos protegidos

**[X] Executar como Administrador pode causar problemas:**

- Arquivos criados com proprietário "Administrador"
- Problemas de permissões para acessar os arquivos depois
- Word pode não conseguir acessar os templates
- Operação desnecessária e insegura

## [*] O que o Script Faz

### 1. Validação Pré-instalação
- [OK] Verifica versão do Windows (10+)
- [OK] Verifica versão do PowerShell (5.1+)
- [OK] Confirma acesso à rede corporativa
- [OK] Testa permissões de escrita no perfil do usuário
- [OK] Valida existência dos arquivos de origem

### 2. Backup Automático
- [LOCK] Renomeia pasta Templates existente com timestamp
- [LOCK] Mantém histórico dos últimos 5 backups
- [LOCK] Remove backups antigos automaticamente
- [LOCK] Formato: `Templates_backup_YYYYMMDD_HHMMSS`

### 3. Instalação
- [DIR] Copia `stamp.png` para `%USERPROFILE%\chainsaw\assets\`
- [DIR] Copia Templates para `%APPDATA%\Microsoft\Templates\`
- [DIR] Preserva toda estrutura de pastas e arquivos
- [DIR] Verifica integridade dos arquivos copiados

### 4. Sistema de Log
- [LOG] Registra todas as operações
- [LOG] Salva em `%USERPROFILE%\chainsaw\logs\`
- [LOG] Formato: `install_YYYYMMDD_HHMMSS.log`
- [LOG] Inclui timestamps, níveis e mensagens detalhadas

### 5. Tratamento de Erros
- [SEC] Validação completa antes de iniciar
- [SEC] Rollback automático em caso de falha
- [SEC] Mensagens de erro claras e acionáveis
- [SEC] Não interrompe em avisos não críticos

## [LOCK] Bypass Automático de Execução (Novo!)

O script agora possui um **mecanismo de auto-relançamento seguro** que elimina a necessidade de configurar manualmente a política de execução do PowerShell.

### Como Funciona

1. **Detecção Automática**: O script detecta se a política de execução impede sua execução
2. **Informação Clara**: Exibe informações de segurança sobre o que será feito
3. **Relançamento Seguro**: Relança-se automaticamente com `-ExecutionPolicy Bypass`
4. **Temporário**: O bypass é válido APENAS para esta execução do script
5. **Sem Alterações**: A política do sistema permanece inalterada
6. **Sem Admin**: Nenhum privilégio de administrador é necessário

### Garantias de Segurança

[OK] **Isolado**: Apenas este script específico é executado com bypass  
[OK] **Temporário**: O bypass expira automaticamente quando o script termina  
[OK] **Transparente**: Todas as ações são informadas ao usuário  
[OK] **Auditável**: Tudo é registrado no arquivo de log  
[OK] **Sem Admin**: Não requer nem usa privilégios elevados  
[OK] **Reversível**: A política original permanece intacta  

### Uso

Simplesmente execute o script normalmente a partir do perfil do usuário:

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

Se necessário, o script se relançará automaticamente. Você verá:

```
[LOCK] Verificando política de execução...
   Política atual (CurrentUser): Restricted
[!]  Política de execução restritiva detectada.
[SYNC] Relançando script com bypass temporário...

[i]  SEGURANÇA:
   • Apenas ESTE script será executado com bypass
   • A política do sistema NÃO será alterada
   • O bypass expira quando o script terminar
   • Nenhum privilégio de administrador é usado

[OK] Executando com bypass temporário (seguro)
```

## [LOCK][>>] Como Usar

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
cd "$env:USERPROFILE\chainsaw"
.\test-permissions.ps1
```

Este script verifica:
- [OK] Se você NÃO está executando como administrador
- [OK] Permissões de escrita em `%USERPROFILE%`
- [OK] Permissões de escrita em `%APPDATA%`
- [OK] Capacidade de criar, renomear e copiar arquivos/pastas

### Teste Rápido de Instalação (Opcional)

Para verificar pré-requisitos sem modificar nada:

```powershell
cd "$env:USERPROFILE\chainsaw"
.\test-install.ps1
```

### Instalação Interativa (Padrão)

**Método Recomendado - Usando o Launcher Seguro:**

```cmd
cd "%USERPROFILE%\chainsaw"
install.cmd
```

**Alternativa - Execução Direta do PowerShell:**

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

O script irá:

1. **Verificar e ajustar automaticamente a política de execução** (bypass temporário seguro)
2. Verificar pré-requisitos
3. Mostrar o que será feito
4. Pedir confirmação
5. Executar a instalação
6. Exibir resultado detalhado

[LOCK] **Segurança do Bypass Automático:**

- [OK] Apenas ESTE script é executado com bypass
- [OK] A política do sistema NÃO é alterada permanentemente
- [OK] O bypass expira automaticamente quando o script termina
- [OK] Nenhum privilégio de administrador é necessário ou usado
- [OK] Totalmente transparente e seguro
- [OK] O launcher `.cmd` funciona em QUALQUER política de execução

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

[!] **Não recomendado** - Instala sem criar backup:

```cmd
install.cmd -NoBackup
```

### Instalação com Caminho Customizado

Se os arquivos estiverem em outro local:

```cmd
install.cmd -SourcePath "C:\outro\caminho\chainsaw"
```

## [INFO] Exemplo de Execução

```
╔════════════════════════════════════════════════════════════════╗
║          CHAINSAW - Instalação de Configurações do Word       ║
╚════════════════════════════════════════════════════════════════╝

[LOG] Arquivo de log: C:\Users\csantos\chainsaw\logs\install_20251105_143022.log

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ETAPA 1: Verificação de Pré-requisitos
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[i] Verificando pré-requisitos...
[OK] Sistema operacional: Windows 10.0 [OK]
[OK] PowerShell versão: 5.1.19041.4894 [OK]
[i] Verificando arquivos de origem: C:\Users\csantos\chainsaw
[OK] Arquivos de origem encontrados [OK]
[OK] Permissões de escrita no perfil do usuário confirmadas [OK]

[... mais output ...]

╔════════════════════════════════════════════════════════════════╗
║              INSTALAÇÃO CONCLUÍDA COM SUCESSO!                 ║
╚════════════════════════════════════════════════════════════════╝

[CHART] Resumo da Instalação:
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
Caminho de Origem: C:\Users\csantos\chainsaw
================================================================================

[2025-11-05 14:30:22] [INFO] === INÍCIO DA INSTALAÇÃO ===
[2025-11-05 14:30:22] [INFO] Verificando pré-requisitos...
[2025-11-05 14:30:22] [SUCCESS] Sistema operacional: Windows 10.0 [OK]
[2025-11-05 14:30:23] [SUCCESS] PowerShell versão: 5.1.19041.4894 [OK]
[2025-11-05 14:30:23] [INFO] Verificando acesso ao caminho de rede: ...
[2025-11-05 14:30:24] [SUCCESS] Acesso ao caminho de rede confirmado [OK]
...
```

## [TOOL] Solução de Problemas

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

[!] **Nota:** A solução automática é mais segura, pois não altera permanentemente as configurações do sistema.

### Erro: Arquivos de origem não encontrados

**Problema:** "Arquivos de origem não encontrados" ou "Não foi possível acessar o caminho"

**Possíveis causas:**
1. Pasta `CHAINSAW` não está no perfil do usuário
2. Arquivos `stamp.png` ou pasta `Templates` ausentes
3. Caminho incorreto especificado

**Solução:**
1. Verifique se a pasta está em: `%USERPROFILE%\chainsaw`
2. Certifique-se que os arquivos necessários estão presentes:
   - `assets\stamp.png`
   - `configs\Templates\`
3. Se os arquivos estão em outro local, use: `install.cmd -SourcePath "C:\caminho\correto"`

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

## [SEC] Segurança

### O que o script NÃO faz

- [X] Não requer privilégios de administrador
- [X] Não modifica arquivos do sistema
- [X] Não altera registro do Windows
- [X] Não instala software adicional
- [X] Não faz comunicação externa
- [X] Não coleta dados do usuário

### O que o script faz para segurança

- [OK] Valida todos os inputs
- [OK] Cria backup antes de modificar
- [OK] Registra todas as operações em log
- [OK] Reverte mudanças em caso de erro
- [OK] Verifica integridade dos arquivos
- [OK] Opera apenas no perfil do usuário

## [DIR] Estrutura de Arquivos Criada

Após a instalação, a seguinte estrutura será criada:

```
%USERPROFILE%\
├─ CHAINSAW\
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

## [SYNC] Atualizações

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
