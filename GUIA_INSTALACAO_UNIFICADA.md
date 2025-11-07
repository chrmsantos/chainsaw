# Guia de Instalação Unificada - CHAINSAW

## [*] Visão Geral

A partir da **versão 2.0.0**, o processo de instalação do CHAINSAW foi unificado em um único script que:

- [OK] Instala templates do Word
- [OK] Detecta e importa personalizações automaticamente (se disponíveis)
- [OK] Cria backups de segurança
- [OK] Registra todas as operações em log

## [INFO] Pré-requisitos

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Microsoft Word fechado durante a instalação
- Pasta `chainsaw` no perfil do usuário (%USERPROFILE%)

## [>>] Instalação Rápida

### Passo 1: Copiar Arquivos

Copie a Pasta `chainsaw` completa para:
```
 # Guia de Instalação Unificada - CHAINSAW

## [*] Visão Geral

A partir da **versão 2.0.0**, o processo de instalação do CHAINSAW foi unificado em um único script que:

- [OK] Instala templates do Word
- [OK] Detecta e importa personalizações automaticamente (se disponíveis)
- [OK] Cria backups de segurança
- [OK] Registra todas as operações em log

## [INFO] Pré-requisitos

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Microsoft Word fechado durante a instalação
- Pasta `chainsaw` no perfil do usuário (%USERPROFILE%)

## [>>] Instalação Rápida

### Passo 1: Copiar Arquivos

Copie a Pasta `chainsaw` completa para:
```
C:\Users\[seu_usuario]\CHAINSAW
```

### Passo 2: Executar Instalação

**Método Recomendado** (funciona com qualquer política de execução):

1. Abra o Explorador de Arquivos
2. Navegue até `C:\Users\[seu_usuario]\CHAINSAW`
3. Dê um duplo-clique em: **`install.cmd`**

**Método Alternativo** (via PowerShell):

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

### Passo 3: Acompanhar Processo

O instalador executará automaticamente:

```
ETAPA 1: Verificação de Pré-requisitos     [OK]
ETAPA 2: Validação de Arquivos             [OK]
ETAPA 3: Cópia do Arquivo de Imagem        [OK]
ETAPA 4: Backup da Pasta Templates         [OK]
ETAPA 5: Cópia da Pasta Templates          [OK]
ETAPA 6: Importação de Personalizações     [OK] (se disponível)
```

## [ART] Importação Automática de Personalizações

### Como Funciona

Se a pasta `exported-config` for detectada dentro da Pasta `chainsaw`, o instalador:

1. **Detecta** automaticamente as personalizações exportadas
2. **Informa** o conteúdo que será importado:
   - Faixa de Opções Personalizada (Ribbon)
   - Partes Rápidas (Quick Parts)
   - Blocos de Construção (Building Blocks)
   - Temas de Documentos
   - Template Normal.dotm

3. **Solicita confirmação** (a menos que use `-Force`)
4. **Cria backup** das personalizações atuais
5. **Importa** todas as personalizações

### Estrutura Esperada

```
C:\Users\[usuario]\chainsaw\
├── install.ps1
├── install.cmd
├── export-config.ps1
├── export-config.cmd
├── configs/
│   └── Templates/
│       └── (arquivos de templates)
├── exported-config/              ← Pasta de personalizações (opcional)
│   ├── MANIFEST.json
│   ├── Templates/
│   │   └── Normal.dotm
│   ├── RibbonCustomization/
│   ├── OfficeCustomUI/
│   └── (outros arquivos)
└── assets/
    └── stamp.png
```

## [CFG] Opções de Instalação

### Instalação Padrão (Interativa)

```cmd
install.cmd
```

- Solicita confirmação para cada etapa importante
- Cria backups automáticos
- Importa personalizações (se disponíveis e confirmado)

### Instalação Automática (Modo Force)

```cmd
install.cmd -Force
```

- Não solicita confirmações
- Executa todas as etapas automaticamente
- Ideal para scripts de deploy

### Instalar Apenas Templates (Sem Personalizações)

```cmd
install.cmd -SkipCustomizations
```

- Instala apenas os templates
- Ignora a pasta `exported-config` mesmo se existir

### Sem Backup (Não Recomendado)

```cmd
install.cmd -NoBackup
```

- Não cria backup das configurações existentes
- Use apenas se tiver certeza do que está fazendo

### Combinando Opções

```cmd
install.cmd -Force -SkipCustomizations
```

## [PKG] Exportar Personalizações (Máquina de Origem)

Para transferir suas personalizações do Word para outra máquina:

### 1. Exportar na Máquina de Origem

```cmd
export-config.cmd
```

Isso criará a pasta `exported-config` com todas as suas personalizações.

### 2. Transferir para Máquina de Destino

Copie a Pasta `chainsaw` completa (incluindo `exported-config`) para a máquina de destino:

```
Origem:  C:\Users\[usuario_origem]\chainsaw\
         └── exported-config\  (gerado pelo export)

Destino: C:\Users\[usuario_destino]\chainsaw\
         └── exported-config\  (copiado da origem)
```

### 3. Instalar na Máquina de Destino

```cmd
install.cmd
```

O instalador detectará automaticamente a pasta `exported-config` e oferecerá importar as personalizações.

## [CHART] Logs e Diagnósticos

Todos os logs são salvos em:
```
%USERPROFILE%\chainsaw\logs\
├── install_20240115_143022.log
├── export_20240115_142100.log
└── (outros logs)
```

### Verificar Último Log

```powershell
notepad "$env:USERPROFILE\chainsaw\logs\$(Get-ChildItem $env:USERPROFILE\chainsaw\logs\install_*.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name)"
```

## [?] Perguntas Frequentes

### O que acontece se eu executar install.cmd sem exported-config?

O instalador funciona normalmente, instalando apenas os templates. Nenhum erro ocorrerá.

### Posso executar install.cmd múltiplas vezes?

Sim! Cada execução cria um novo backup com timestamp. Os 5 backups mais recentes são mantidos automaticamente.

### Como saber se as personalizações foram importadas?

1. Verifique o log em `%USERPROFILE%\chainsaw\logs\`
2. Abra o Word e verifique suas abas personalizadas na Faixa de Opções
3. Procure por "ETAPA 6: Importação de Personalizações" na saída do instalador

### Posso importar personalizações depois?

Sim! Se você pulou a importação durante a instalação inicial, basta:

1. Obter a pasta `exported-config` 
2. Colocá-la em `C:\Users\[usuario]\chainsaw\`
3. Executar `install.cmd` novamente

### O que são os arquivos .cmd?

São launchers seguros que:
- Executam os scripts PowerShell com bypass temporário
- Funcionam em qualquer política de execução
- Não alteram configurações permanentes do sistema
- São mais fáceis de usar (duplo-clique)

## [SEC] Segurança

### Bypass de Política de Execução

Os arquivos `.cmd` usam bypass temporário:

```cmd
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "script.ps1"
```

**Isso é seguro?** [OK] SIM

- Apenas o script especificado é executado
- Não há alteração permanente nas políticas do sistema
- O bypass expira quando o script termina
- Nenhum privilégio de administrador é necessário

### Backups Automáticos

Antes de qualquer modificação:
- [OK] Templates atuais → `Templates_backup_[timestamp]`
- [OK] Personalizações atuais → `chainsaw\backups\word-customizations_[timestamp]`

Para restaurar um backup manualmente:
```powershell
# Templates
$backup = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates_backup_20240115_143022"
$target = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates"
Remove-Item $target -Recurse -Force
Rename-Item $backup "Templates"
```

## 📚 Documentação Adicional

- **[README.md](README.md)** - Visão geral completa do sistema
- **[INSTALL.md](INSTALL.md)** - Instruções detalhadas de instalação
- **[GUIA_RAPIDO_EXPORT_IMPORT.md](GUIA_RAPIDO_EXPORT_IMPORT.md)** - Guia de exportação/importação

## [NEW] Mudanças da Versão 2.0.0

### O que mudou?

**Antes (v1.x):**
```
1. install.ps1      → Instala templates
2. import-config.ps1 → Importa personalizações (manual)
```

**Agora (v2.0+):**
```
1. install.ps1      → Instala templates + importa personalizações automaticamente
```

### Scripts Removidos

Os seguintes scripts foram consolidados ou removidos:
- [X] `import-config.ps1` (funcionalidade integrada ao `install.ps1`)
- [X] `import-config.cmd` (não é mais necessário)
- [X] `start-install.ps1` (substituído por `install.cmd`)
- [X] `test-*.ps1` (scripts de teste legados)

### Scripts Mantidos

- [OK] `install.ps1` - **Instalador unificado** (agora com importação integrada)
- [OK] `install.cmd` - Launcher seguro
- [OK] `export-config.ps1` - Exportação de personalizações
- [OK] `export-config.cmd` - Launcher seguro para exportação

## [ART] Exemplos de Uso

### Cenário 1: Instalação Nova (Sem Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates instalados, nenhuma personalização importada.

### Cenário 2: Instalação com Personalizações

```cmd
# Na máquina de origem
cd %USERPROFILE%\chainsaw
export-config.cmd

# Copiar pasta CHAINSAW completa para máquina de destino

# Na máquina de destino
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates + personalizações instalados.

### Cenário 3: Atualização de Templates (Preservar Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -SkipCustomizations
```

Resultado: Apenas templates atualizados, personalizações não tocadas.

### Cenário 4: Deploy Automatizado

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -Force
```

Resultado: Instalação completamente automática, sem interação do usuário.

## [?] Troubleshooting

### Erro: "Script não pode ser carregado"

**Solução:** Use `install.cmd` ao invés de `install.ps1` diretamente.

### Erro: "Acesso negado" ao copiar arquivos

**Solução:** 
1. Feche o Microsoft Word
2. Verifique se não está executando como Administrador (não use "Executar como administrador")
3. Execute novamente

### Personalizações não aparecem no Word

**Solução:**
1. Feche completamente o Word
2. Abra o Gerenciador de Tarefas e encerre qualquer processo `WINWORD.EXE`
3. Abra o Word novamente

### Quero ver o que será importado antes de confirmar

**Solução:** Execute em modo interativo (sem `-Force`) e revise a lista apresentada antes de confirmar.

---

## [i] Dica Final

Para uma instalação mais rápida e fácil:

1. [OK] Use `install.cmd` (duplo-clique)
2. [OK] Mantenha a pasta `exported-config` se tiver personalizações
3. [OK] Feche o Word antes de instalar
4. [OK] Não execute como Administrador

**É isso! A instalação ficou muito mais simples.** [NEW]
.Value -replace 'CHAINSAW', 'chainsaw' 
```

### Passo 2: Executar Instalação

**Método Recomendado** (funciona com qualquer política de execução):

1. Abra o Explorador de Arquivos
2. Navegue até ` # Guia de Instalação Unificada - CHAINSAW

## [*] Visão Geral

A partir da **versão 2.0.0**, o processo de instalação do CHAINSAW foi unificado em um único script que:

- [OK] Instala templates do Word
- [OK] Detecta e importa personalizações automaticamente (se disponíveis)
- [OK] Cria backups de segurança
- [OK] Registra todas as operações em log

## [INFO] Pré-requisitos

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Microsoft Word fechado durante a instalação
- Pasta `chainsaw` no perfil do usuário (%USERPROFILE%)

## [>>] Instalação Rápida

### Passo 1: Copiar Arquivos

Copie a Pasta `chainsaw` completa para:
```
C:\Users\[seu_usuario]\CHAINSAW
```

### Passo 2: Executar Instalação

**Método Recomendado** (funciona com qualquer política de execução):

1. Abra o Explorador de Arquivos
2. Navegue até `C:\Users\[seu_usuario]\CHAINSAW`
3. Dê um duplo-clique em: **`install.cmd`**

**Método Alternativo** (via PowerShell):

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

### Passo 3: Acompanhar Processo

O instalador executará automaticamente:

```
ETAPA 1: Verificação de Pré-requisitos     [OK]
ETAPA 2: Validação de Arquivos             [OK]
ETAPA 3: Cópia do Arquivo de Imagem        [OK]
ETAPA 4: Backup da Pasta Templates         [OK]
ETAPA 5: Cópia da Pasta Templates          [OK]
ETAPA 6: Importação de Personalizações     [OK] (se disponível)
```

## [ART] Importação Automática de Personalizações

### Como Funciona

Se a pasta `exported-config` for detectada dentro da Pasta `chainsaw`, o instalador:

1. **Detecta** automaticamente as personalizações exportadas
2. **Informa** o conteúdo que será importado:
   - Faixa de Opções Personalizada (Ribbon)
   - Partes Rápidas (Quick Parts)
   - Blocos de Construção (Building Blocks)
   - Temas de Documentos
   - Template Normal.dotm

3. **Solicita confirmação** (a menos que use `-Force`)
4. **Cria backup** das personalizações atuais
5. **Importa** todas as personalizações

### Estrutura Esperada

```
C:\Users\[usuario]\chainsaw\
├── install.ps1
├── install.cmd
├── export-config.ps1
├── export-config.cmd
├── configs/
│   └── Templates/
│       └── (arquivos de templates)
├── exported-config/              ← Pasta de personalizações (opcional)
│   ├── MANIFEST.json
│   ├── Templates/
│   │   └── Normal.dotm
│   ├── RibbonCustomization/
│   ├── OfficeCustomUI/
│   └── (outros arquivos)
└── assets/
    └── stamp.png
```

## [CFG] Opções de Instalação

### Instalação Padrão (Interativa)

```cmd
install.cmd
```

- Solicita confirmação para cada etapa importante
- Cria backups automáticos
- Importa personalizações (se disponíveis e confirmado)

### Instalação Automática (Modo Force)

```cmd
install.cmd -Force
```

- Não solicita confirmações
- Executa todas as etapas automaticamente
- Ideal para scripts de deploy

### Instalar Apenas Templates (Sem Personalizações)

```cmd
install.cmd -SkipCustomizations
```

- Instala apenas os templates
- Ignora a pasta `exported-config` mesmo se existir

### Sem Backup (Não Recomendado)

```cmd
install.cmd -NoBackup
```

- Não cria backup das configurações existentes
- Use apenas se tiver certeza do que está fazendo

### Combinando Opções

```cmd
install.cmd -Force -SkipCustomizations
```

## [PKG] Exportar Personalizações (Máquina de Origem)

Para transferir suas personalizações do Word para outra máquina:

### 1. Exportar na Máquina de Origem

```cmd
export-config.cmd
```

Isso criará a pasta `exported-config` com todas as suas personalizações.

### 2. Transferir para Máquina de Destino

Copie a Pasta `chainsaw` completa (incluindo `exported-config`) para a máquina de destino:

```
Origem:  C:\Users\[usuario_origem]\chainsaw\
         └── exported-config\  (gerado pelo export)

Destino: C:\Users\[usuario_destino]\chainsaw\
         └── exported-config\  (copiado da origem)
```

### 3. Instalar na Máquina de Destino

```cmd
install.cmd
```

O instalador detectará automaticamente a pasta `exported-config` e oferecerá importar as personalizações.

## [CHART] Logs e Diagnósticos

Todos os logs são salvos em:
```
%USERPROFILE%\chainsaw\logs\
├── install_20240115_143022.log
├── export_20240115_142100.log
└── (outros logs)
```

### Verificar Último Log

```powershell
notepad "$env:USERPROFILE\chainsaw\logs\$(Get-ChildItem $env:USERPROFILE\chainsaw\logs\install_*.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name)"
```

## [?] Perguntas Frequentes

### O que acontece se eu executar install.cmd sem exported-config?

O instalador funciona normalmente, instalando apenas os templates. Nenhum erro ocorrerá.

### Posso executar install.cmd múltiplas vezes?

Sim! Cada execução cria um novo backup com timestamp. Os 5 backups mais recentes são mantidos automaticamente.

### Como saber se as personalizações foram importadas?

1. Verifique o log em `%USERPROFILE%\chainsaw\logs\`
2. Abra o Word e verifique suas abas personalizadas na Faixa de Opções
3. Procure por "ETAPA 6: Importação de Personalizações" na saída do instalador

### Posso importar personalizações depois?

Sim! Se você pulou a importação durante a instalação inicial, basta:

1. Obter a pasta `exported-config` 
2. Colocá-la em `C:\Users\[usuario]\chainsaw\`
3. Executar `install.cmd` novamente

### O que são os arquivos .cmd?

São launchers seguros que:
- Executam os scripts PowerShell com bypass temporário
- Funcionam em qualquer política de execução
- Não alteram configurações permanentes do sistema
- São mais fáceis de usar (duplo-clique)

## [SEC] Segurança

### Bypass de Política de Execução

Os arquivos `.cmd` usam bypass temporário:

```cmd
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "script.ps1"
```

**Isso é seguro?** [OK] SIM

- Apenas o script especificado é executado
- Não há alteração permanente nas políticas do sistema
- O bypass expira quando o script termina
- Nenhum privilégio de administrador é necessário

### Backups Automáticos

Antes de qualquer modificação:
- [OK] Templates atuais → `Templates_backup_[timestamp]`
- [OK] Personalizações atuais → `chainsaw\backups\word-customizations_[timestamp]`

Para restaurar um backup manualmente:
```powershell
# Templates
$backup = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates_backup_20240115_143022"
$target = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates"
Remove-Item $target -Recurse -Force
Rename-Item $backup "Templates"
```

## 📚 Documentação Adicional

- **[README.md](README.md)** - Visão geral completa do sistema
- **[INSTALL.md](INSTALL.md)** - Instruções detalhadas de instalação
- **[GUIA_RAPIDO_EXPORT_IMPORT.md](GUIA_RAPIDO_EXPORT_IMPORT.md)** - Guia de exportação/importação

## [NEW] Mudanças da Versão 2.0.0

### O que mudou?

**Antes (v1.x):**
```
1. install.ps1      → Instala templates
2. import-config.ps1 → Importa personalizações (manual)
```

**Agora (v2.0+):**
```
1. install.ps1      → Instala templates + importa personalizações automaticamente
```

### Scripts Removidos

Os seguintes scripts foram consolidados ou removidos:
- [X] `import-config.ps1` (funcionalidade integrada ao `install.ps1`)
- [X] `import-config.cmd` (não é mais necessário)
- [X] `start-install.ps1` (substituído por `install.cmd`)
- [X] `test-*.ps1` (scripts de teste legados)

### Scripts Mantidos

- [OK] `install.ps1` - **Instalador unificado** (agora com importação integrada)
- [OK] `install.cmd` - Launcher seguro
- [OK] `export-config.ps1` - Exportação de personalizações
- [OK] `export-config.cmd` - Launcher seguro para exportação

## [ART] Exemplos de Uso

### Cenário 1: Instalação Nova (Sem Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates instalados, nenhuma personalização importada.

### Cenário 2: Instalação com Personalizações

```cmd
# Na máquina de origem
cd %USERPROFILE%\chainsaw
export-config.cmd

# Copiar pasta CHAINSAW completa para máquina de destino

# Na máquina de destino
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates + personalizações instalados.

### Cenário 3: Atualização de Templates (Preservar Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -SkipCustomizations
```

Resultado: Apenas templates atualizados, personalizações não tocadas.

### Cenário 4: Deploy Automatizado

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -Force
```

Resultado: Instalação completamente automática, sem interação do usuário.

## [?] Troubleshooting

### Erro: "Script não pode ser carregado"

**Solução:** Use `install.cmd` ao invés de `install.ps1` diretamente.

### Erro: "Acesso negado" ao copiar arquivos

**Solução:** 
1. Feche o Microsoft Word
2. Verifique se não está executando como Administrador (não use "Executar como administrador")
3. Execute novamente

### Personalizações não aparecem no Word

**Solução:**
1. Feche completamente o Word
2. Abra o Gerenciador de Tarefas e encerre qualquer processo `WINWORD.EXE`
3. Abra o Word novamente

### Quero ver o que será importado antes de confirmar

**Solução:** Execute em modo interativo (sem `-Force`) e revise a lista apresentada antes de confirmar.

---

## [i] Dica Final

Para uma instalação mais rápida e fácil:

1. [OK] Use `install.cmd` (duplo-clique)
2. [OK] Mantenha a pasta `exported-config` se tiver personalizações
3. [OK] Feche o Word antes de instalar
4. [OK] Não execute como Administrador

**É isso! A instalação ficou muito mais simples.** [NEW]
.Value -replace 'CHAINSAW', 'chainsaw' `
3. Dê um duplo-clique em: **`install.cmd`**

**Método Alternativo** (via PowerShell):

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

### Passo 3: Acompanhar Processo

O instalador executará automaticamente:

```
ETAPA 1: Verificação de Pré-requisitos     [OK]
ETAPA 2: Validação de Arquivos             [OK]
ETAPA 3: Cópia do Arquivo de Imagem        [OK]
ETAPA 4: Backup da Pasta Templates         [OK]
ETAPA 5: Cópia da Pasta Templates          [OK]
ETAPA 6: Importação de Personalizações     [OK] (se disponível)
```

## [ART] Importação Automática de Personalizações

### Como Funciona

Se a pasta `exported-config` for detectada dentro da Pasta `chainsaw`, o instalador:

1. **Detecta** automaticamente as personalizações exportadas
2. **Informa** o conteúdo que será importado:
   - Faixa de Opções Personalizada (Ribbon)
   - Partes Rápidas (Quick Parts)
   - Blocos de Construção (Building Blocks)
   - Temas de Documentos
   - Template Normal.dotm

3. **Solicita confirmação** (a menos que use `-Force`)
4. **Cria backup** das personalizações atuais
5. **Importa** todas as personalizações

### Estrutura Esperada

```
 # Guia de Instalação Unificada - CHAINSAW

## [*] Visão Geral

A partir da **versão 2.0.0**, o processo de instalação do CHAINSAW foi unificado em um único script que:

- [OK] Instala templates do Word
- [OK] Detecta e importa personalizações automaticamente (se disponíveis)
- [OK] Cria backups de segurança
- [OK] Registra todas as operações em log

## [INFO] Pré-requisitos

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Microsoft Word fechado durante a instalação
- Pasta `chainsaw` no perfil do usuário (%USERPROFILE%)

## [>>] Instalação Rápida

### Passo 1: Copiar Arquivos

Copie a Pasta `chainsaw` completa para:
```
C:\Users\[seu_usuario]\CHAINSAW
```

### Passo 2: Executar Instalação

**Método Recomendado** (funciona com qualquer política de execução):

1. Abra o Explorador de Arquivos
2. Navegue até `C:\Users\[seu_usuario]\CHAINSAW`
3. Dê um duplo-clique em: **`install.cmd`**

**Método Alternativo** (via PowerShell):

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

### Passo 3: Acompanhar Processo

O instalador executará automaticamente:

```
ETAPA 1: Verificação de Pré-requisitos     [OK]
ETAPA 2: Validação de Arquivos             [OK]
ETAPA 3: Cópia do Arquivo de Imagem        [OK]
ETAPA 4: Backup da Pasta Templates         [OK]
ETAPA 5: Cópia da Pasta Templates          [OK]
ETAPA 6: Importação de Personalizações     [OK] (se disponível)
```

## [ART] Importação Automática de Personalizações

### Como Funciona

Se a pasta `exported-config` for detectada dentro da Pasta `chainsaw`, o instalador:

1. **Detecta** automaticamente as personalizações exportadas
2. **Informa** o conteúdo que será importado:
   - Faixa de Opções Personalizada (Ribbon)
   - Partes Rápidas (Quick Parts)
   - Blocos de Construção (Building Blocks)
   - Temas de Documentos
   - Template Normal.dotm

3. **Solicita confirmação** (a menos que use `-Force`)
4. **Cria backup** das personalizações atuais
5. **Importa** todas as personalizações

### Estrutura Esperada

```
C:\Users\[usuario]\chainsaw\
├── install.ps1
├── install.cmd
├── export-config.ps1
├── export-config.cmd
├── configs/
│   └── Templates/
│       └── (arquivos de templates)
├── exported-config/              ← Pasta de personalizações (opcional)
│   ├── MANIFEST.json
│   ├── Templates/
│   │   └── Normal.dotm
│   ├── RibbonCustomization/
│   ├── OfficeCustomUI/
│   └── (outros arquivos)
└── assets/
    └── stamp.png
```

## [CFG] Opções de Instalação

### Instalação Padrão (Interativa)

```cmd
install.cmd
```

- Solicita confirmação para cada etapa importante
- Cria backups automáticos
- Importa personalizações (se disponíveis e confirmado)

### Instalação Automática (Modo Force)

```cmd
install.cmd -Force
```

- Não solicita confirmações
- Executa todas as etapas automaticamente
- Ideal para scripts de deploy

### Instalar Apenas Templates (Sem Personalizações)

```cmd
install.cmd -SkipCustomizations
```

- Instala apenas os templates
- Ignora a pasta `exported-config` mesmo se existir

### Sem Backup (Não Recomendado)

```cmd
install.cmd -NoBackup
```

- Não cria backup das configurações existentes
- Use apenas se tiver certeza do que está fazendo

### Combinando Opções

```cmd
install.cmd -Force -SkipCustomizations
```

## [PKG] Exportar Personalizações (Máquina de Origem)

Para transferir suas personalizações do Word para outra máquina:

### 1. Exportar na Máquina de Origem

```cmd
export-config.cmd
```

Isso criará a pasta `exported-config` com todas as suas personalizações.

### 2. Transferir para Máquina de Destino

Copie a Pasta `chainsaw` completa (incluindo `exported-config`) para a máquina de destino:

```
Origem:  C:\Users\[usuario_origem]\chainsaw\
         └── exported-config\  (gerado pelo export)

Destino: C:\Users\[usuario_destino]\chainsaw\
         └── exported-config\  (copiado da origem)
```

### 3. Instalar na Máquina de Destino

```cmd
install.cmd
```

O instalador detectará automaticamente a pasta `exported-config` e oferecerá importar as personalizações.

## [CHART] Logs e Diagnósticos

Todos os logs são salvos em:
```
%USERPROFILE%\chainsaw\logs\
├── install_20240115_143022.log
├── export_20240115_142100.log
└── (outros logs)
```

### Verificar Último Log

```powershell
notepad "$env:USERPROFILE\chainsaw\logs\$(Get-ChildItem $env:USERPROFILE\chainsaw\logs\install_*.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name)"
```

## [?] Perguntas Frequentes

### O que acontece se eu executar install.cmd sem exported-config?

O instalador funciona normalmente, instalando apenas os templates. Nenhum erro ocorrerá.

### Posso executar install.cmd múltiplas vezes?

Sim! Cada execução cria um novo backup com timestamp. Os 5 backups mais recentes são mantidos automaticamente.

### Como saber se as personalizações foram importadas?

1. Verifique o log em `%USERPROFILE%\chainsaw\logs\`
2. Abra o Word e verifique suas abas personalizadas na Faixa de Opções
3. Procure por "ETAPA 6: Importação de Personalizações" na saída do instalador

### Posso importar personalizações depois?

Sim! Se você pulou a importação durante a instalação inicial, basta:

1. Obter a pasta `exported-config` 
2. Colocá-la em `C:\Users\[usuario]\chainsaw\`
3. Executar `install.cmd` novamente

### O que são os arquivos .cmd?

São launchers seguros que:
- Executam os scripts PowerShell com bypass temporário
- Funcionam em qualquer política de execução
- Não alteram configurações permanentes do sistema
- São mais fáceis de usar (duplo-clique)

## [SEC] Segurança

### Bypass de Política de Execução

Os arquivos `.cmd` usam bypass temporário:

```cmd
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "script.ps1"
```

**Isso é seguro?** [OK] SIM

- Apenas o script especificado é executado
- Não há alteração permanente nas políticas do sistema
- O bypass expira quando o script termina
- Nenhum privilégio de administrador é necessário

### Backups Automáticos

Antes de qualquer modificação:
- [OK] Templates atuais → `Templates_backup_[timestamp]`
- [OK] Personalizações atuais → `chainsaw\backups\word-customizations_[timestamp]`

Para restaurar um backup manualmente:
```powershell
# Templates
$backup = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates_backup_20240115_143022"
$target = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates"
Remove-Item $target -Recurse -Force
Rename-Item $backup "Templates"
```

## 📚 Documentação Adicional

- **[README.md](README.md)** - Visão geral completa do sistema
- **[INSTALL.md](INSTALL.md)** - Instruções detalhadas de instalação
- **[GUIA_RAPIDO_EXPORT_IMPORT.md](GUIA_RAPIDO_EXPORT_IMPORT.md)** - Guia de exportação/importação

## [NEW] Mudanças da Versão 2.0.0

### O que mudou?

**Antes (v1.x):**
```
1. install.ps1      → Instala templates
2. import-config.ps1 → Importa personalizações (manual)
```

**Agora (v2.0+):**
```
1. install.ps1      → Instala templates + importa personalizações automaticamente
```

### Scripts Removidos

Os seguintes scripts foram consolidados ou removidos:
- [X] `import-config.ps1` (funcionalidade integrada ao `install.ps1`)
- [X] `import-config.cmd` (não é mais necessário)
- [X] `start-install.ps1` (substituído por `install.cmd`)
- [X] `test-*.ps1` (scripts de teste legados)

### Scripts Mantidos

- [OK] `install.ps1` - **Instalador unificado** (agora com importação integrada)
- [OK] `install.cmd` - Launcher seguro
- [OK] `export-config.ps1` - Exportação de personalizações
- [OK] `export-config.cmd` - Launcher seguro para exportação

## [ART] Exemplos de Uso

### Cenário 1: Instalação Nova (Sem Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates instalados, nenhuma personalização importada.

### Cenário 2: Instalação com Personalizações

```cmd
# Na máquina de origem
cd %USERPROFILE%\chainsaw
export-config.cmd

# Copiar pasta CHAINSAW completa para máquina de destino

# Na máquina de destino
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates + personalizações instalados.

### Cenário 3: Atualização de Templates (Preservar Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -SkipCustomizations
```

Resultado: Apenas templates atualizados, personalizações não tocadas.

### Cenário 4: Deploy Automatizado

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -Force
```

Resultado: Instalação completamente automática, sem interação do usuário.

## [?] Troubleshooting

### Erro: "Script não pode ser carregado"

**Solução:** Use `install.cmd` ao invés de `install.ps1` diretamente.

### Erro: "Acesso negado" ao copiar arquivos

**Solução:** 
1. Feche o Microsoft Word
2. Verifique se não está executando como Administrador (não use "Executar como administrador")
3. Execute novamente

### Personalizações não aparecem no Word

**Solução:**
1. Feche completamente o Word
2. Abra o Gerenciador de Tarefas e encerre qualquer processo `WINWORD.EXE`
3. Abra o Word novamente

### Quero ver o que será importado antes de confirmar

**Solução:** Execute em modo interativo (sem `-Force`) e revise a lista apresentada antes de confirmar.

---

## [i] Dica Final

Para uma instalação mais rápida e fácil:

1. [OK] Use `install.cmd` (duplo-clique)
2. [OK] Mantenha a pasta `exported-config` se tiver personalizações
3. [OK] Feche o Word antes de instalar
4. [OK] Não execute como Administrador

**É isso! A instalação ficou muito mais simples.** [NEW]
.Value -replace 'CHAINSAW', 'chainsaw' \
├── install.ps1
├── install.cmd
├── export-config.ps1
├── export-config.cmd
├── configs/
│   └── Templates/
│       └── (arquivos de templates)
├── exported-config/              ← Pasta de personalizações (opcional)
│   ├── MANIFEST.json
│   ├── Templates/
│   │   └── Normal.dotm
│   ├── RibbonCustomization/
│   ├── OfficeCustomUI/
│   └── (outros arquivos)
└── assets/
    └── stamp.png
```

## [CFG] Opções de Instalação

### Instalação Padrão (Interativa)

```cmd
install.cmd
```

- Solicita confirmação para cada etapa importante
- Cria backups automáticos
- Importa personalizações (se disponíveis e confirmado)

### Instalação Automática (Modo Force)

```cmd
install.cmd -Force
```

- Não solicita confirmações
- Executa todas as etapas automaticamente
- Ideal para scripts de deploy

### Instalar Apenas Templates (Sem Personalizações)

```cmd
install.cmd -SkipCustomizations
```

- Instala apenas os templates
- Ignora a pasta `exported-config` mesmo se existir

### Sem Backup (Não Recomendado)

```cmd
install.cmd -NoBackup
```

- Não cria backup das configurações existentes
- Use apenas se tiver certeza do que está fazendo

### Combinando Opções

```cmd
install.cmd -Force -SkipCustomizations
```

## [PKG] Exportar Personalizações (Máquina de Origem)

Para transferir suas personalizações do Word para outra máquina:

### 1. Exportar na Máquina de Origem

```cmd
export-config.cmd
```

Isso criará a pasta `exported-config` com todas as suas personalizações.

### 2. Transferir para Máquina de Destino

Copie a Pasta `chainsaw` completa (incluindo `exported-config`) para a máquina de destino:

```
Origem:   # Guia de Instalação Unificada - CHAINSAW

## [*] Visão Geral

A partir da **versão 2.0.0**, o processo de instalação do CHAINSAW foi unificado em um único script que:

- [OK] Instala templates do Word
- [OK] Detecta e importa personalizações automaticamente (se disponíveis)
- [OK] Cria backups de segurança
- [OK] Registra todas as operações em log

## [INFO] Pré-requisitos

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Microsoft Word fechado durante a instalação
- Pasta `chainsaw` no perfil do usuário (%USERPROFILE%)

## [>>] Instalação Rápida

### Passo 1: Copiar Arquivos

Copie a Pasta `chainsaw` completa para:
```
C:\Users\[seu_usuario]\CHAINSAW
```

### Passo 2: Executar Instalação

**Método Recomendado** (funciona com qualquer política de execução):

1. Abra o Explorador de Arquivos
2. Navegue até `C:\Users\[seu_usuario]\CHAINSAW`
3. Dê um duplo-clique em: **`install.cmd`**

**Método Alternativo** (via PowerShell):

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

### Passo 3: Acompanhar Processo

O instalador executará automaticamente:

```
ETAPA 1: Verificação de Pré-requisitos     [OK]
ETAPA 2: Validação de Arquivos             [OK]
ETAPA 3: Cópia do Arquivo de Imagem        [OK]
ETAPA 4: Backup da Pasta Templates         [OK]
ETAPA 5: Cópia da Pasta Templates          [OK]
ETAPA 6: Importação de Personalizações     [OK] (se disponível)
```

## [ART] Importação Automática de Personalizações

### Como Funciona

Se a pasta `exported-config` for detectada dentro da Pasta `chainsaw`, o instalador:

1. **Detecta** automaticamente as personalizações exportadas
2. **Informa** o conteúdo que será importado:
   - Faixa de Opções Personalizada (Ribbon)
   - Partes Rápidas (Quick Parts)
   - Blocos de Construção (Building Blocks)
   - Temas de Documentos
   - Template Normal.dotm

3. **Solicita confirmação** (a menos que use `-Force`)
4. **Cria backup** das personalizações atuais
5. **Importa** todas as personalizações

### Estrutura Esperada

```
C:\Users\[usuario]\chainsaw\
├── install.ps1
├── install.cmd
├── export-config.ps1
├── export-config.cmd
├── configs/
│   └── Templates/
│       └── (arquivos de templates)
├── exported-config/              ← Pasta de personalizações (opcional)
│   ├── MANIFEST.json
│   ├── Templates/
│   │   └── Normal.dotm
│   ├── RibbonCustomization/
│   ├── OfficeCustomUI/
│   └── (outros arquivos)
└── assets/
    └── stamp.png
```

## [CFG] Opções de Instalação

### Instalação Padrão (Interativa)

```cmd
install.cmd
```

- Solicita confirmação para cada etapa importante
- Cria backups automáticos
- Importa personalizações (se disponíveis e confirmado)

### Instalação Automática (Modo Force)

```cmd
install.cmd -Force
```

- Não solicita confirmações
- Executa todas as etapas automaticamente
- Ideal para scripts de deploy

### Instalar Apenas Templates (Sem Personalizações)

```cmd
install.cmd -SkipCustomizations
```

- Instala apenas os templates
- Ignora a pasta `exported-config` mesmo se existir

### Sem Backup (Não Recomendado)

```cmd
install.cmd -NoBackup
```

- Não cria backup das configurações existentes
- Use apenas se tiver certeza do que está fazendo

### Combinando Opções

```cmd
install.cmd -Force -SkipCustomizations
```

## [PKG] Exportar Personalizações (Máquina de Origem)

Para transferir suas personalizações do Word para outra máquina:

### 1. Exportar na Máquina de Origem

```cmd
export-config.cmd
```

Isso criará a pasta `exported-config` com todas as suas personalizações.

### 2. Transferir para Máquina de Destino

Copie a Pasta `chainsaw` completa (incluindo `exported-config`) para a máquina de destino:

```
Origem:  C:\Users\[usuario_origem]\chainsaw\
         └── exported-config\  (gerado pelo export)

Destino: C:\Users\[usuario_destino]\chainsaw\
         └── exported-config\  (copiado da origem)
```

### 3. Instalar na Máquina de Destino

```cmd
install.cmd
```

O instalador detectará automaticamente a pasta `exported-config` e oferecerá importar as personalizações.

## [CHART] Logs e Diagnósticos

Todos os logs são salvos em:
```
%USERPROFILE%\chainsaw\logs\
├── install_20240115_143022.log
├── export_20240115_142100.log
└── (outros logs)
```

### Verificar Último Log

```powershell
notepad "$env:USERPROFILE\chainsaw\logs\$(Get-ChildItem $env:USERPROFILE\chainsaw\logs\install_*.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name)"
```

## [?] Perguntas Frequentes

### O que acontece se eu executar install.cmd sem exported-config?

O instalador funciona normalmente, instalando apenas os templates. Nenhum erro ocorrerá.

### Posso executar install.cmd múltiplas vezes?

Sim! Cada execução cria um novo backup com timestamp. Os 5 backups mais recentes são mantidos automaticamente.

### Como saber se as personalizações foram importadas?

1. Verifique o log em `%USERPROFILE%\chainsaw\logs\`
2. Abra o Word e verifique suas abas personalizadas na Faixa de Opções
3. Procure por "ETAPA 6: Importação de Personalizações" na saída do instalador

### Posso importar personalizações depois?

Sim! Se você pulou a importação durante a instalação inicial, basta:

1. Obter a pasta `exported-config` 
2. Colocá-la em `C:\Users\[usuario]\chainsaw\`
3. Executar `install.cmd` novamente

### O que são os arquivos .cmd?

São launchers seguros que:
- Executam os scripts PowerShell com bypass temporário
- Funcionam em qualquer política de execução
- Não alteram configurações permanentes do sistema
- São mais fáceis de usar (duplo-clique)

## [SEC] Segurança

### Bypass de Política de Execução

Os arquivos `.cmd` usam bypass temporário:

```cmd
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "script.ps1"
```

**Isso é seguro?** [OK] SIM

- Apenas o script especificado é executado
- Não há alteração permanente nas políticas do sistema
- O bypass expira quando o script termina
- Nenhum privilégio de administrador é necessário

### Backups Automáticos

Antes de qualquer modificação:
- [OK] Templates atuais → `Templates_backup_[timestamp]`
- [OK] Personalizações atuais → `chainsaw\backups\word-customizations_[timestamp]`

Para restaurar um backup manualmente:
```powershell
# Templates
$backup = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates_backup_20240115_143022"
$target = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates"
Remove-Item $target -Recurse -Force
Rename-Item $backup "Templates"
```

## 📚 Documentação Adicional

- **[README.md](README.md)** - Visão geral completa do sistema
- **[INSTALL.md](INSTALL.md)** - Instruções detalhadas de instalação
- **[GUIA_RAPIDO_EXPORT_IMPORT.md](GUIA_RAPIDO_EXPORT_IMPORT.md)** - Guia de exportação/importação

## [NEW] Mudanças da Versão 2.0.0

### O que mudou?

**Antes (v1.x):**
```
1. install.ps1      → Instala templates
2. import-config.ps1 → Importa personalizações (manual)
```

**Agora (v2.0+):**
```
1. install.ps1      → Instala templates + importa personalizações automaticamente
```

### Scripts Removidos

Os seguintes scripts foram consolidados ou removidos:
- [X] `import-config.ps1` (funcionalidade integrada ao `install.ps1`)
- [X] `import-config.cmd` (não é mais necessário)
- [X] `start-install.ps1` (substituído por `install.cmd`)
- [X] `test-*.ps1` (scripts de teste legados)

### Scripts Mantidos

- [OK] `install.ps1` - **Instalador unificado** (agora com importação integrada)
- [OK] `install.cmd` - Launcher seguro
- [OK] `export-config.ps1` - Exportação de personalizações
- [OK] `export-config.cmd` - Launcher seguro para exportação

## [ART] Exemplos de Uso

### Cenário 1: Instalação Nova (Sem Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates instalados, nenhuma personalização importada.

### Cenário 2: Instalação com Personalizações

```cmd
# Na máquina de origem
cd %USERPROFILE%\chainsaw
export-config.cmd

# Copiar pasta CHAINSAW completa para máquina de destino

# Na máquina de destino
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates + personalizações instalados.

### Cenário 3: Atualização de Templates (Preservar Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -SkipCustomizations
```

Resultado: Apenas templates atualizados, personalizações não tocadas.

### Cenário 4: Deploy Automatizado

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -Force
```

Resultado: Instalação completamente automática, sem interação do usuário.

## [?] Troubleshooting

### Erro: "Script não pode ser carregado"

**Solução:** Use `install.cmd` ao invés de `install.ps1` diretamente.

### Erro: "Acesso negado" ao copiar arquivos

**Solução:** 
1. Feche o Microsoft Word
2. Verifique se não está executando como Administrador (não use "Executar como administrador")
3. Execute novamente

### Personalizações não aparecem no Word

**Solução:**
1. Feche completamente o Word
2. Abra o Gerenciador de Tarefas e encerre qualquer processo `WINWORD.EXE`
3. Abra o Word novamente

### Quero ver o que será importado antes de confirmar

**Solução:** Execute em modo interativo (sem `-Force`) e revise a lista apresentada antes de confirmar.

---

## [i] Dica Final

Para uma instalação mais rápida e fácil:

1. [OK] Use `install.cmd` (duplo-clique)
2. [OK] Mantenha a pasta `exported-config` se tiver personalizações
3. [OK] Feche o Word antes de instalar
4. [OK] Não execute como Administrador

**É isso! A instalação ficou muito mais simples.** [NEW]
.Value -replace 'CHAINSAW', 'chainsaw' \
         └── exported-config\  (gerado pelo export)

Destino:  # Guia de Instalação Unificada - CHAINSAW

## [*] Visão Geral

A partir da **versão 2.0.0**, o processo de instalação do CHAINSAW foi unificado em um único script que:

- [OK] Instala templates do Word
- [OK] Detecta e importa personalizações automaticamente (se disponíveis)
- [OK] Cria backups de segurança
- [OK] Registra todas as operações em log

## [INFO] Pré-requisitos

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Microsoft Word fechado durante a instalação
- Pasta `chainsaw` no perfil do usuário (%USERPROFILE%)

## [>>] Instalação Rápida

### Passo 1: Copiar Arquivos

Copie a Pasta `chainsaw` completa para:
```
C:\Users\[seu_usuario]\CHAINSAW
```

### Passo 2: Executar Instalação

**Método Recomendado** (funciona com qualquer política de execução):

1. Abra o Explorador de Arquivos
2. Navegue até `C:\Users\[seu_usuario]\CHAINSAW`
3. Dê um duplo-clique em: **`install.cmd`**

**Método Alternativo** (via PowerShell):

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

### Passo 3: Acompanhar Processo

O instalador executará automaticamente:

```
ETAPA 1: Verificação de Pré-requisitos     [OK]
ETAPA 2: Validação de Arquivos             [OK]
ETAPA 3: Cópia do Arquivo de Imagem        [OK]
ETAPA 4: Backup da Pasta Templates         [OK]
ETAPA 5: Cópia da Pasta Templates          [OK]
ETAPA 6: Importação de Personalizações     [OK] (se disponível)
```

## [ART] Importação Automática de Personalizações

### Como Funciona

Se a pasta `exported-config` for detectada dentro da Pasta `chainsaw`, o instalador:

1. **Detecta** automaticamente as personalizações exportadas
2. **Informa** o conteúdo que será importado:
   - Faixa de Opções Personalizada (Ribbon)
   - Partes Rápidas (Quick Parts)
   - Blocos de Construção (Building Blocks)
   - Temas de Documentos
   - Template Normal.dotm

3. **Solicita confirmação** (a menos que use `-Force`)
4. **Cria backup** das personalizações atuais
5. **Importa** todas as personalizações

### Estrutura Esperada

```
C:\Users\[usuario]\chainsaw\
├── install.ps1
├── install.cmd
├── export-config.ps1
├── export-config.cmd
├── configs/
│   └── Templates/
│       └── (arquivos de templates)
├── exported-config/              ← Pasta de personalizações (opcional)
│   ├── MANIFEST.json
│   ├── Templates/
│   │   └── Normal.dotm
│   ├── RibbonCustomization/
│   ├── OfficeCustomUI/
│   └── (outros arquivos)
└── assets/
    └── stamp.png
```

## [CFG] Opções de Instalação

### Instalação Padrão (Interativa)

```cmd
install.cmd
```

- Solicita confirmação para cada etapa importante
- Cria backups automáticos
- Importa personalizações (se disponíveis e confirmado)

### Instalação Automática (Modo Force)

```cmd
install.cmd -Force
```

- Não solicita confirmações
- Executa todas as etapas automaticamente
- Ideal para scripts de deploy

### Instalar Apenas Templates (Sem Personalizações)

```cmd
install.cmd -SkipCustomizations
```

- Instala apenas os templates
- Ignora a pasta `exported-config` mesmo se existir

### Sem Backup (Não Recomendado)

```cmd
install.cmd -NoBackup
```

- Não cria backup das configurações existentes
- Use apenas se tiver certeza do que está fazendo

### Combinando Opções

```cmd
install.cmd -Force -SkipCustomizations
```

## [PKG] Exportar Personalizações (Máquina de Origem)

Para transferir suas personalizações do Word para outra máquina:

### 1. Exportar na Máquina de Origem

```cmd
export-config.cmd
```

Isso criará a pasta `exported-config` com todas as suas personalizações.

### 2. Transferir para Máquina de Destino

Copie a Pasta `chainsaw` completa (incluindo `exported-config`) para a máquina de destino:

```
Origem:  C:\Users\[usuario_origem]\chainsaw\
         └── exported-config\  (gerado pelo export)

Destino: C:\Users\[usuario_destino]\chainsaw\
         └── exported-config\  (copiado da origem)
```

### 3. Instalar na Máquina de Destino

```cmd
install.cmd
```

O instalador detectará automaticamente a pasta `exported-config` e oferecerá importar as personalizações.

## [CHART] Logs e Diagnósticos

Todos os logs são salvos em:
```
%USERPROFILE%\chainsaw\logs\
├── install_20240115_143022.log
├── export_20240115_142100.log
└── (outros logs)
```

### Verificar Último Log

```powershell
notepad "$env:USERPROFILE\chainsaw\logs\$(Get-ChildItem $env:USERPROFILE\chainsaw\logs\install_*.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name)"
```

## [?] Perguntas Frequentes

### O que acontece se eu executar install.cmd sem exported-config?

O instalador funciona normalmente, instalando apenas os templates. Nenhum erro ocorrerá.

### Posso executar install.cmd múltiplas vezes?

Sim! Cada execução cria um novo backup com timestamp. Os 5 backups mais recentes são mantidos automaticamente.

### Como saber se as personalizações foram importadas?

1. Verifique o log em `%USERPROFILE%\chainsaw\logs\`
2. Abra o Word e verifique suas abas personalizadas na Faixa de Opções
3. Procure por "ETAPA 6: Importação de Personalizações" na saída do instalador

### Posso importar personalizações depois?

Sim! Se você pulou a importação durante a instalação inicial, basta:

1. Obter a pasta `exported-config` 
2. Colocá-la em `C:\Users\[usuario]\chainsaw\`
3. Executar `install.cmd` novamente

### O que são os arquivos .cmd?

São launchers seguros que:
- Executam os scripts PowerShell com bypass temporário
- Funcionam em qualquer política de execução
- Não alteram configurações permanentes do sistema
- São mais fáceis de usar (duplo-clique)

## [SEC] Segurança

### Bypass de Política de Execução

Os arquivos `.cmd` usam bypass temporário:

```cmd
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "script.ps1"
```

**Isso é seguro?** [OK] SIM

- Apenas o script especificado é executado
- Não há alteração permanente nas políticas do sistema
- O bypass expira quando o script termina
- Nenhum privilégio de administrador é necessário

### Backups Automáticos

Antes de qualquer modificação:
- [OK] Templates atuais → `Templates_backup_[timestamp]`
- [OK] Personalizações atuais → `chainsaw\backups\word-customizations_[timestamp]`

Para restaurar um backup manualmente:
```powershell
# Templates
$backup = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates_backup_20240115_143022"
$target = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates"
Remove-Item $target -Recurse -Force
Rename-Item $backup "Templates"
```

## 📚 Documentação Adicional

- **[README.md](README.md)** - Visão geral completa do sistema
- **[INSTALL.md](INSTALL.md)** - Instruções detalhadas de instalação
- **[GUIA_RAPIDO_EXPORT_IMPORT.md](GUIA_RAPIDO_EXPORT_IMPORT.md)** - Guia de exportação/importação

## [NEW] Mudanças da Versão 2.0.0

### O que mudou?

**Antes (v1.x):**
```
1. install.ps1      → Instala templates
2. import-config.ps1 → Importa personalizações (manual)
```

**Agora (v2.0+):**
```
1. install.ps1      → Instala templates + importa personalizações automaticamente
```

### Scripts Removidos

Os seguintes scripts foram consolidados ou removidos:
- [X] `import-config.ps1` (funcionalidade integrada ao `install.ps1`)
- [X] `import-config.cmd` (não é mais necessário)
- [X] `start-install.ps1` (substituído por `install.cmd`)
- [X] `test-*.ps1` (scripts de teste legados)

### Scripts Mantidos

- [OK] `install.ps1` - **Instalador unificado** (agora com importação integrada)
- [OK] `install.cmd` - Launcher seguro
- [OK] `export-config.ps1` - Exportação de personalizações
- [OK] `export-config.cmd` - Launcher seguro para exportação

## [ART] Exemplos de Uso

### Cenário 1: Instalação Nova (Sem Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates instalados, nenhuma personalização importada.

### Cenário 2: Instalação com Personalizações

```cmd
# Na máquina de origem
cd %USERPROFILE%\chainsaw
export-config.cmd

# Copiar pasta CHAINSAW completa para máquina de destino

# Na máquina de destino
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates + personalizações instalados.

### Cenário 3: Atualização de Templates (Preservar Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -SkipCustomizations
```

Resultado: Apenas templates atualizados, personalizações não tocadas.

### Cenário 4: Deploy Automatizado

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -Force
```

Resultado: Instalação completamente automática, sem interação do usuário.

## [?] Troubleshooting

### Erro: "Script não pode ser carregado"

**Solução:** Use `install.cmd` ao invés de `install.ps1` diretamente.

### Erro: "Acesso negado" ao copiar arquivos

**Solução:** 
1. Feche o Microsoft Word
2. Verifique se não está executando como Administrador (não use "Executar como administrador")
3. Execute novamente

### Personalizações não aparecem no Word

**Solução:**
1. Feche completamente o Word
2. Abra o Gerenciador de Tarefas e encerre qualquer processo `WINWORD.EXE`
3. Abra o Word novamente

### Quero ver o que será importado antes de confirmar

**Solução:** Execute em modo interativo (sem `-Force`) e revise a lista apresentada antes de confirmar.

---

## [i] Dica Final

Para uma instalação mais rápida e fácil:

1. [OK] Use `install.cmd` (duplo-clique)
2. [OK] Mantenha a pasta `exported-config` se tiver personalizações
3. [OK] Feche o Word antes de instalar
4. [OK] Não execute como Administrador

**É isso! A instalação ficou muito mais simples.** [NEW]
.Value -replace 'CHAINSAW', 'chainsaw' \
         └── exported-config\  (copiado da origem)
```

### 3. Instalar na Máquina de Destino

```cmd
install.cmd
```

O instalador detectará automaticamente a pasta `exported-config` e oferecerá importar as personalizações.

## [CHART] Logs e Diagnósticos

Todos os logs são salvos em:
```
%USERPROFILE%\chainsaw\logs\
├── install_20240115_143022.log
├── export_20240115_142100.log
└── (outros logs)
```

### Verificar Último Log

```powershell
notepad "$env:USERPROFILE\chainsaw\logs\$(Get-ChildItem $env:USERPROFILE\chainsaw\logs\install_*.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name)"
```

## [?] Perguntas Frequentes

### O que acontece se eu executar install.cmd sem exported-config?

O instalador funciona normalmente, instalando apenas os templates. Nenhum erro ocorrerá.

### Posso executar install.cmd múltiplas vezes?

Sim! Cada execução cria um novo backup com timestamp. Os 5 backups mais recentes são mantidos automaticamente.

### Como saber se as personalizações foram importadas?

1. Verifique o log em `%USERPROFILE%\chainsaw\logs\`
2. Abra o Word e verifique suas abas personalizadas na Faixa de Opções
3. Procure por "ETAPA 6: Importação de Personalizações" na saída do instalador

### Posso importar personalizações depois?

Sim! Se você pulou a importação durante a instalação inicial, basta:

1. Obter a pasta `exported-config` 
2. Colocá-la em ` # Guia de Instalação Unificada - CHAINSAW

## [*] Visão Geral

A partir da **versão 2.0.0**, o processo de instalação do CHAINSAW foi unificado em um único script que:

- [OK] Instala templates do Word
- [OK] Detecta e importa personalizações automaticamente (se disponíveis)
- [OK] Cria backups de segurança
- [OK] Registra todas as operações em log

## [INFO] Pré-requisitos

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Microsoft Word fechado durante a instalação
- Pasta `chainsaw` no perfil do usuário (%USERPROFILE%)

## [>>] Instalação Rápida

### Passo 1: Copiar Arquivos

Copie a Pasta `chainsaw` completa para:
```
C:\Users\[seu_usuario]\CHAINSAW
```

### Passo 2: Executar Instalação

**Método Recomendado** (funciona com qualquer política de execução):

1. Abra o Explorador de Arquivos
2. Navegue até `C:\Users\[seu_usuario]\CHAINSAW`
3. Dê um duplo-clique em: **`install.cmd`**

**Método Alternativo** (via PowerShell):

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

### Passo 3: Acompanhar Processo

O instalador executará automaticamente:

```
ETAPA 1: Verificação de Pré-requisitos     [OK]
ETAPA 2: Validação de Arquivos             [OK]
ETAPA 3: Cópia do Arquivo de Imagem        [OK]
ETAPA 4: Backup da Pasta Templates         [OK]
ETAPA 5: Cópia da Pasta Templates          [OK]
ETAPA 6: Importação de Personalizações     [OK] (se disponível)
```

## [ART] Importação Automática de Personalizações

### Como Funciona

Se a pasta `exported-config` for detectada dentro da Pasta `chainsaw`, o instalador:

1. **Detecta** automaticamente as personalizações exportadas
2. **Informa** o conteúdo que será importado:
   - Faixa de Opções Personalizada (Ribbon)
   - Partes Rápidas (Quick Parts)
   - Blocos de Construção (Building Blocks)
   - Temas de Documentos
   - Template Normal.dotm

3. **Solicita confirmação** (a menos que use `-Force`)
4. **Cria backup** das personalizações atuais
5. **Importa** todas as personalizações

### Estrutura Esperada

```
C:\Users\[usuario]\chainsaw\
├── install.ps1
├── install.cmd
├── export-config.ps1
├── export-config.cmd
├── configs/
│   └── Templates/
│       └── (arquivos de templates)
├── exported-config/              ← Pasta de personalizações (opcional)
│   ├── MANIFEST.json
│   ├── Templates/
│   │   └── Normal.dotm
│   ├── RibbonCustomization/
│   ├── OfficeCustomUI/
│   └── (outros arquivos)
└── assets/
    └── stamp.png
```

## [CFG] Opções de Instalação

### Instalação Padrão (Interativa)

```cmd
install.cmd
```

- Solicita confirmação para cada etapa importante
- Cria backups automáticos
- Importa personalizações (se disponíveis e confirmado)

### Instalação Automática (Modo Force)

```cmd
install.cmd -Force
```

- Não solicita confirmações
- Executa todas as etapas automaticamente
- Ideal para scripts de deploy

### Instalar Apenas Templates (Sem Personalizações)

```cmd
install.cmd -SkipCustomizations
```

- Instala apenas os templates
- Ignora a pasta `exported-config` mesmo se existir

### Sem Backup (Não Recomendado)

```cmd
install.cmd -NoBackup
```

- Não cria backup das configurações existentes
- Use apenas se tiver certeza do que está fazendo

### Combinando Opções

```cmd
install.cmd -Force -SkipCustomizations
```

## [PKG] Exportar Personalizações (Máquina de Origem)

Para transferir suas personalizações do Word para outra máquina:

### 1. Exportar na Máquina de Origem

```cmd
export-config.cmd
```

Isso criará a pasta `exported-config` com todas as suas personalizações.

### 2. Transferir para Máquina de Destino

Copie a Pasta `chainsaw` completa (incluindo `exported-config`) para a máquina de destino:

```
Origem:  C:\Users\[usuario_origem]\chainsaw\
         └── exported-config\  (gerado pelo export)

Destino: C:\Users\[usuario_destino]\chainsaw\
         └── exported-config\  (copiado da origem)
```

### 3. Instalar na Máquina de Destino

```cmd
install.cmd
```

O instalador detectará automaticamente a pasta `exported-config` e oferecerá importar as personalizações.

## [CHART] Logs e Diagnósticos

Todos os logs são salvos em:
```
%USERPROFILE%\chainsaw\logs\
├── install_20240115_143022.log
├── export_20240115_142100.log
└── (outros logs)
```

### Verificar Último Log

```powershell
notepad "$env:USERPROFILE\chainsaw\logs\$(Get-ChildItem $env:USERPROFILE\chainsaw\logs\install_*.log | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty Name)"
```

## [?] Perguntas Frequentes

### O que acontece se eu executar install.cmd sem exported-config?

O instalador funciona normalmente, instalando apenas os templates. Nenhum erro ocorrerá.

### Posso executar install.cmd múltiplas vezes?

Sim! Cada execução cria um novo backup com timestamp. Os 5 backups mais recentes são mantidos automaticamente.

### Como saber se as personalizações foram importadas?

1. Verifique o log em `%USERPROFILE%\chainsaw\logs\`
2. Abra o Word e verifique suas abas personalizadas na Faixa de Opções
3. Procure por "ETAPA 6: Importação de Personalizações" na saída do instalador

### Posso importar personalizações depois?

Sim! Se você pulou a importação durante a instalação inicial, basta:

1. Obter a pasta `exported-config` 
2. Colocá-la em `C:\Users\[usuario]\chainsaw\`
3. Executar `install.cmd` novamente

### O que são os arquivos .cmd?

São launchers seguros que:
- Executam os scripts PowerShell com bypass temporário
- Funcionam em qualquer política de execução
- Não alteram configurações permanentes do sistema
- São mais fáceis de usar (duplo-clique)

## [SEC] Segurança

### Bypass de Política de Execução

Os arquivos `.cmd` usam bypass temporário:

```cmd
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "script.ps1"
```

**Isso é seguro?** [OK] SIM

- Apenas o script especificado é executado
- Não há alteração permanente nas políticas do sistema
- O bypass expira quando o script termina
- Nenhum privilégio de administrador é necessário

### Backups Automáticos

Antes de qualquer modificação:
- [OK] Templates atuais → `Templates_backup_[timestamp]`
- [OK] Personalizações atuais → `chainsaw\backups\word-customizations_[timestamp]`

Para restaurar um backup manualmente:
```powershell
# Templates
$backup = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates_backup_20240115_143022"
$target = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates"
Remove-Item $target -Recurse -Force
Rename-Item $backup "Templates"
```

## 📚 Documentação Adicional

- **[README.md](README.md)** - Visão geral completa do sistema
- **[INSTALL.md](INSTALL.md)** - Instruções detalhadas de instalação
- **[GUIA_RAPIDO_EXPORT_IMPORT.md](GUIA_RAPIDO_EXPORT_IMPORT.md)** - Guia de exportação/importação

## [NEW] Mudanças da Versão 2.0.0

### O que mudou?

**Antes (v1.x):**
```
1. install.ps1      → Instala templates
2. import-config.ps1 → Importa personalizações (manual)
```

**Agora (v2.0+):**
```
1. install.ps1      → Instala templates + importa personalizações automaticamente
```

### Scripts Removidos

Os seguintes scripts foram consolidados ou removidos:
- [X] `import-config.ps1` (funcionalidade integrada ao `install.ps1`)
- [X] `import-config.cmd` (não é mais necessário)
- [X] `start-install.ps1` (substituído por `install.cmd`)
- [X] `test-*.ps1` (scripts de teste legados)

### Scripts Mantidos

- [OK] `install.ps1` - **Instalador unificado** (agora com importação integrada)
- [OK] `install.cmd` - Launcher seguro
- [OK] `export-config.ps1` - Exportação de personalizações
- [OK] `export-config.cmd` - Launcher seguro para exportação

## [ART] Exemplos de Uso

### Cenário 1: Instalação Nova (Sem Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates instalados, nenhuma personalização importada.

### Cenário 2: Instalação com Personalizações

```cmd
# Na máquina de origem
cd %USERPROFILE%\chainsaw
export-config.cmd

# Copiar pasta CHAINSAW completa para máquina de destino

# Na máquina de destino
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates + personalizações instalados.

### Cenário 3: Atualização de Templates (Preservar Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -SkipCustomizations
```

Resultado: Apenas templates atualizados, personalizações não tocadas.

### Cenário 4: Deploy Automatizado

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -Force
```

Resultado: Instalação completamente automática, sem interação do usuário.

## [?] Troubleshooting

### Erro: "Script não pode ser carregado"

**Solução:** Use `install.cmd` ao invés de `install.ps1` diretamente.

### Erro: "Acesso negado" ao copiar arquivos

**Solução:** 
1. Feche o Microsoft Word
2. Verifique se não está executando como Administrador (não use "Executar como administrador")
3. Execute novamente

### Personalizações não aparecem no Word

**Solução:**
1. Feche completamente o Word
2. Abra o Gerenciador de Tarefas e encerre qualquer processo `WINWORD.EXE`
3. Abra o Word novamente

### Quero ver o que será importado antes de confirmar

**Solução:** Execute em modo interativo (sem `-Force`) e revise a lista apresentada antes de confirmar.

---

## [i] Dica Final

Para uma instalação mais rápida e fácil:

1. [OK] Use `install.cmd` (duplo-clique)
2. [OK] Mantenha a pasta `exported-config` se tiver personalizações
3. [OK] Feche o Word antes de instalar
4. [OK] Não execute como Administrador

**É isso! A instalação ficou muito mais simples.** [NEW]
.Value -replace 'CHAINSAW', 'chainsaw' \`
3. Executar `install.cmd` novamente

### O que são os arquivos .cmd?

São launchers seguros que:
- Executam os scripts PowerShell com bypass temporário
- Funcionam em qualquer política de execução
- Não alteram configurações permanentes do sistema
- São mais fáceis de usar (duplo-clique)

## [SEC] Segurança

### Bypass de Política de Execução

Os arquivos `.cmd` usam bypass temporário:

```cmd
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "script.ps1"
```

**Isso é seguro?** [OK] SIM

- Apenas o script especificado é executado
- Não há alteração permanente nas políticas do sistema
- O bypass expira quando o script termina
- Nenhum privilégio de administrador é necessário

### Backups Automáticos

Antes de qualquer modificação:
- [OK] Templates atuais → `Templates_backup_[timestamp]`
- [OK] Personalizações atuais → `chainsaw\backups\word-customizations_[timestamp]`

Para restaurar um backup manualmente:
```powershell
# Templates
$backup = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates_backup_20240115_143022"
$target = "C:\Users\[usuario]\AppData\Roaming\Microsoft\Templates"
Remove-Item $target -Recurse -Force
Rename-Item $backup "Templates"
```

## 📚 Documentação Adicional

- **[README.md](README.md)** - Visão geral completa do sistema
- **[INSTALL.md](INSTALL.md)** - Instruções detalhadas de instalação
- **[GUIA_RAPIDO_EXPORT_IMPORT.md](GUIA_RAPIDO_EXPORT_IMPORT.md)** - Guia de exportação/importação

## [NEW] Mudanças da Versão 2.0.0

### O que mudou?

**Antes (v1.x):**
```
1. install.ps1      → Instala templates
2. import-config.ps1 → Importa personalizações (manual)
```

**Agora (v2.0+):**
```
1. install.ps1      → Instala templates + importa personalizações automaticamente
```

### Scripts Removidos

Os seguintes scripts foram consolidados ou removidos:
- [X] `import-config.ps1` (funcionalidade integrada ao `install.ps1`)
- [X] `import-config.cmd` (não é mais necessário)
- [X] `start-install.ps1` (substituído por `install.cmd`)
- [X] `test-*.ps1` (scripts de teste legados)

### Scripts Mantidos

- [OK] `install.ps1` - **Instalador unificado** (agora com importação integrada)
- [OK] `install.cmd` - Launcher seguro
- [OK] `export-config.ps1` - Exportação de personalizações
- [OK] `export-config.cmd` - Launcher seguro para exportação

## [ART] Exemplos de Uso

### Cenário 1: Instalação Nova (Sem Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates instalados, nenhuma personalização importada.

### Cenário 2: Instalação com Personalizações

```cmd
# Na máquina de origem
cd %USERPROFILE%\chainsaw
export-config.cmd

# Copiar pasta CHAINSAW completa para máquina de destino

# Na máquina de destino
cd %USERPROFILE%\chainsaw
install.cmd
```

Resultado: Templates + personalizações instalados.

### Cenário 3: Atualização de Templates (Preservar Personalizações)

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -SkipCustomizations
```

Resultado: Apenas templates atualizados, personalizações não tocadas.

### Cenário 4: Deploy Automatizado

```cmd
cd %USERPROFILE%\chainsaw
install.cmd -Force
```

Resultado: Instalação completamente automática, sem interação do usuário.

## [?] Troubleshooting

### Erro: "Script não pode ser carregado"

**Solução:** Use `install.cmd` ao invés de `install.ps1` diretamente.

### Erro: "Acesso negado" ao copiar arquivos

**Solução:** 
1. Feche o Microsoft Word
2. Verifique se não está executando como Administrador (não use "Executar como administrador")
3. Execute novamente

### Personalizações não aparecem no Word

**Solução:**
1. Feche completamente o Word
2. Abra o Gerenciador de Tarefas e encerre qualquer processo `WINWORD.EXE`
3. Abra o Word novamente

### Quero ver o que será importado antes de confirmar

**Solução:** Execute em modo interativo (sem `-Force`) e revise a lista apresentada antes de confirmar.

---

## [i] Dica Final

Para uma instalação mais rápida e fácil:

1. [OK] Use `install.cmd` (duplo-clique)
2. [OK] Mantenha a pasta `exported-config` se tiver personalizações
3. [OK] Feche o Word antes de instalar
4. [OK] Não execute como Administrador

**É isso! A instalação ficou muito mais simples.** [NEW]
