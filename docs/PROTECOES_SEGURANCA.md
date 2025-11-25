# Proteções de Segurança e Prevenção de Perda de Dados - CHAINSAW

## 🛡️ Visão Geral

Este documento detalha todas as proteções implementadas no sistema CHAINSAW para **prevenir perda acidental de dados** durante processos de instalação e atualização.

## 🚨 Problema Crítico Resolvido

**Situação Anterior**: Era possível que o conteúdo da pasta `chainsaw` fosse deletado sem substituição adequada, causando perda total de dados.

**Solução Implementada**: Sistema robusto de validações, backups obrigatórios e rollback automático.

---

## 📋 Proteções Implementadas

### 1. **Validação de Download (chainsaw_installer.cmd)**

#### 1.1 Validação de Tamanho Mínimo
```batch
REM Verifica tamanho mínimo (arquivo válido deve ter pelo menos 100KB)
if %ZIP_SIZE% LSS 102400 (
    call :Log "[ERRO] Arquivo ZIP muito pequeno (possivelmente corrompido)"
    exit /b 1
)
```

**Proteção**: Previne instalação de arquivos corrompidos ou downloads incompletos.

#### 1.2 Teste de Integridade do ZIP
```batch
REM Testa a integridade do ZIP ANTES de fazer qualquer modificação
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead('%TEMP_ZIP%')
$entryCount = $zip.Entries.Count
```

**Proteção**: Garante que o arquivo ZIP está válido e contém dados reais.

#### 1.3 Validação de Conteúdo Mínimo
```batch
if $entryCount -lt 10 {
    Write-Host "[ERRO] ZIP contem muito poucos arquivos: $entryCount"
    exit 1
}
```

**Proteção**: Previne instalação de ZIPs vazios ou incompletos.

---

### 2. **Backup Obrigatório e Validado**

#### 2.1 Criação de Backup ANTES de Qualquer Modificação
```batch
call :Log "[CRITICO] Criando backup OBRIGATORIO antes de qualquer modificacao..."
xcopy "%INSTALL_DIR%\*" "!BACKUP_DIR!\" /E /H /C /I /Y
```

**Proteção**: Dados originais são preservados antes de qualquer alteração.

#### 2.2 Validação do Backup Criado
```batch
REM VALIDAÇÃO DO BACKUP - CRÍTICO!
set "BACKUP_FILE_COUNT=0"
for /r "!BACKUP_DIR!" %%f in (*) do set /a BACKUP_FILE_COUNT+=1

if %BACKUP_FILE_COUNT% LSS 5 (
    call :Log "[ERRO CRITICO] Backup contem muito poucos arquivos: %BACKUP_FILE_COUNT%"
    call :Log "[ERRO] Instalacao ABORTADA - backup parece incompleto."
    exit /b 1
)
```

**Proteção**: Instalação é **abortada** se o backup falhar ou estiver incompleto.

#### 2.3 Falha de Backup = Instalação Abortada
```batch
if !BACKUP_FAILED! equ 1 (
    call :Log "[ERRO CRITICO] Falha ao criar backup de seguranca!"
    call :Log "[ERRO] NAO E SEGURO CONTINUAR sem backup valido."
    call :Log "[ERRO] Instalacao ABORTADA para proteger seus dados."
    exit /b 1
)
```

**Proteção**: **NUNCA** remove dados existentes sem backup válido.

---

### 3. **Validação de Conteúdo Extraído**

#### 3.1 Validação de Estrutura de Diretórios
```batch
REM Valida presença de pastas essenciais
if not exist "%SOURCE_DIR%\installation" (
    call :Log "[ERRO] Pasta 'installation' nao encontrada no conteudo extraido!"
    set "VALIDATION_FAILED=1"
)

if not exist "%SOURCE_DIR%\installation\inst_scripts\install.cmd" (
    call :Log "[ERRO] Script 'install.cmd' nao encontrado!"
    set "VALIDATION_FAILED=1"
)
```

**Proteção**: Garante que o conteúdo baixado está completo e correto.

#### 3.2 Validação de Quantidade de Arquivos
```batch
set "EXTRACTED_FILE_COUNT=0"
for /r "%SOURCE_DIR%" %%f in (*) do set /a EXTRACTED_FILE_COUNT+=1

if %EXTRACTED_FILE_COUNT% LSS 20 (
    call :Log "[ERRO] Conteudo extraido contem muito poucos arquivos: %EXTRACTED_FILE_COUNT%"
    call :Log "[ERRO] Download pode estar incompleto. Abortando."
    exit /b 1
)
```

**Proteção**: Previne instalação de conteúdo incompleto.

---

### 4. **Operação Atômica (Tudo ou Nada)**

#### 4.1 Extração em Área Temporária
```batch
call :Log "[SEGURANCA] Extraindo para area temporaria primeiro (protecao de dados)..."
set "TEMP_EXTRACT=%TEMP%\chainsaw-extract"
```

**Proteção**: Dados são preparados completamente ANTES de tocar nos arquivos de produção.

#### 4.2 Validação Completa ANTES de Remoção
```batch
REM =============================================================================
REM VALIDAÇÃO CRÍTICA DO CONTEÚDO EXTRAÍDO
REM =============================================================================
call :Log "[CRITICO] Validando conteudo extraido ANTES de instalar..."

REM [... todas as validações ...]

call :Log "[OK] Validacao completa! Seguro para instalar."

REM =============================================================================
REM AGORA SIM: Move os arquivos validados para o destino final
REM =============================================================================
```

**Proteção**: Pasta existente **SOMENTE** é removida APÓS validação completa do novo conteúdo.

---

### 5. **Rollback Automático**

#### 5.1 Detecção de Falha na Cópia
```batch
xcopy "%SOURCE_DIR%\*" "%INSTALL_DIR%\" /E /H /C /I /Y >nul
set "COPY_EXIT=%ERRORLEVEL%"

if %COPY_EXIT% neq 0 (
    call :Log "[ERRO CRITICO] Falha ao copiar arquivos para o destino (erro %COPY_EXIT%)!"
    call :Log "[ROLLBACK] Tentando restaurar backup..."
```

**Proteção**: Falha na cópia dispara rollback imediato.

#### 5.2 Restauração Automática do Backup
```batch
if exist "!BACKUP_DIR!" (
    REM Remove instalação parcial
    if exist "%INSTALL_DIR%" rd /s /q "%INSTALL_DIR%" >nul 2>&1
    
    REM Restaura backup
    xcopy "!BACKUP_DIR!\*" "%INSTALL_DIR%\" /E /H /C /I /Y >nul 2>&1
    if errorlevel 1 (
        call :Log "[ERRO] Falha ao restaurar backup automaticamente!"
        call :Log "[IMPORTANTE] Backup preservado em: !BACKUP_DIR!"
    ) else (
        call :Log "[OK] Backup restaurado com sucesso!"
    )
)
```

**Proteção**: Sistema retorna automaticamente ao estado anterior em caso de falha.

#### 5.3 Validação Final da Instalação
```batch
REM =============================================================================
REM VALIDAÇÃO FINAL DA INSTALAÇÃO
REM =============================================================================
if not exist "%INSTALL_DIR%\installation\inst_scripts\install.cmd" (
    call :Log "[ERRO] install.cmd nao encontrado apos instalacao!"
    set "FINAL_VALIDATION_FAILED=1"
)

if %FINAL_VALIDATION_FAILED% equ 1 (
    call :Log "[ERRO CRITICO] Validacao final FALHOU!"
    call :Log "[ROLLBACK] Restaurando backup..."
    [... restaura backup ...]
)
```

**Proteção**: Instalação é validada após conclusão; rollback se algo estiver errado.

---

### 6. **Validações no install.ps1**

#### 6.1 Validação de Arquivo stamp.png
```powershell
# VALIDAÇÃO CRÍTICA 1: Verifica se arquivo de origem existe
if (-not (Test-Path $SourceFile)) {
    throw "Arquivo stamp.png não encontrado na origem. Instalação abortada."
}

# VALIDAÇÃO CRÍTICA 2: Verifica tamanho mínimo do arquivo
if ($sourceFileInfo.Length -lt 100) {
    throw "Arquivo stamp.png inválido (tamanho suspeito). Instalação abortada."
}

# VALIDAÇÃO CRÍTICA 3: Verifica se o arquivo foi copiado corretamente
if ($sourceSize -ne $destSize) {
    throw "Cópia de stamp.png falhou (tamanhos diferentes). Instalação abortada."
}
```

**Proteção**: Cada arquivo é validado antes, durante e após a cópia.

#### 6.2 Validação de Pasta Templates
```powershell
# VALIDAÇÃO CRÍTICA 1: Verifica se pasta de origem existe
if (-not (Test-Path $SourceFolder)) {
    throw "Pasta Templates não encontrada na origem. Instalação abortada."
}

# VALIDAÇÃO CRÍTICA 2: Verifica se há arquivos na origem
if ($sourceItems.Count -eq 0) {
    throw "Pasta Templates na origem está vazia. Instalação abortada."
}

# VALIDAÇÃO CRÍTICA 3: Verifica se Normal.dotm existe na origem
if (-not (Test-Path $sourceNormalDotm)) {
    throw "Normal.dotm não encontrado na pasta Templates de origem. Instalação abortada."
}

# Valida tamanho mínimo de Normal.dotm
if ($normalDotmSize -lt 10000) {  # Normal.dotm deve ter pelo menos 10KB
    throw "Normal.dotm inválido na origem (tamanho suspeito). Instalação abortada."
}

# VALIDAÇÃO CRÍTICA 4: Verifica que Normal.dotm foi copiado corretamente
if ($destNormalDotmSize -ne $normalDotmSize) {
    throw "Cópia de Normal.dotm falhou (tamanhos diferentes). Instalação abortada."
}
```

**Proteção**: Arquivos críticos são validados em múltiplas etapas.

#### 6.3 Rollback Automático com Validação
```powershell
# VALIDA BACKUP ANTES DE RESTAURAR
$backupItems = Get-ChildItem -Path $backupPath -Recurse -File -ErrorAction Stop
if ($backupItems.Count -eq 0) {
    Write-Log "[ERRO] Backup está vazio - não é seguro restaurar" -Level ERROR
    throw "Backup inválido"
}

Write-Log "Backup validado: $($backupItems.Count) arquivos" -Level INFO

# Restaura backup
Rename-Item -Path $backupPath -NewName "Templates" -Force -ErrorAction Stop

# Valida restauração
$restoredPath = Join-Path $env:APPDATA "Microsoft\Templates"
if (Test-Path $restoredPath) {
    Write-Host "[OK] Backup restaurado com sucesso" -ForegroundColor Green
}
```

**Proteção**: Até o rollback é validado para garantir restauração correta.

---

## 🔍 Cenários Protegidos

| Cenário | Proteção Implementada |
|---------|----------------------|
| **Download corrompido** | Validação de tamanho e integridade do ZIP |
| **Download incompleto** | Validação de quantidade mínima de arquivos |
| **Falha no backup** | Instalação abortada - NUNCA prossegue sem backup |
| **Backup incompleto** | Validação conta arquivos no backup |
| **Conteúdo extraído inválido** | Validação de estrutura de diretórios e arquivos essenciais |
| **Falha na cópia** | Rollback automático restaura estado anterior |
| **Instalação parcial** | Validação final + rollback se necessário |
| **Perda de conexão durante download** | Validação de integridade detecta arquivo corrompido |
| **Disco cheio durante instalação** | Erro na cópia dispara rollback |
| **Arquivo origem corrompido** | Validação de tamanho mínimo e checksums |

---

## 📊 Fluxo de Segurança

```
┌─────────────────────────────────────────────────────┐
│  1. DOWNLOAD                                        │
│     ✓ Validação de tamanho                          │
│     ✓ Teste de integridade do ZIP                   │
│     ✓ Validação de conteúdo mínimo                  │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│  2. EXTRAÇÃO EM ÁREA TEMPORÁRIA                     │
│     ✓ Sem tocar em arquivos de produção             │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│  3. VALIDAÇÃO COMPLETA DO CONTEÚDO                  │
│     ✓ Estrutura de diretórios                       │
│     ✓ Arquivos essenciais presentes                 │
│     ✓ Quantidade mínima de arquivos                 │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│  4. BACKUP OBRIGATÓRIO                              │
│     ✓ Cópia completa de instalação existente        │
│     ✓ Validação do backup                           │
│     ✓ ABORTA se backup falhar                       │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│  5. INSTALAÇÃO                                      │
│     ✓ Remove pasta antiga (backup já validado)      │
│     ✓ Copia novos arquivos                          │
│     ✓ Monitora erros                                │
└─────────────────────────────────────────────────────┘
                         ↓
┌─────────────────────────────────────────────────────┐
│  6. VALIDAÇÃO FINAL                                 │
│     ✓ Verifica arquivos essenciais                  │
│     ✓ Valida integridade                            │
└─────────────────────────────────────────────────────┘
                         ↓
         ┌───────────────┴───────────────┐
         │                               │
    ✓ SUCESSO                      ✗ FALHA
         │                               │
         v                               v
┌─────────────────┐           ┌─────────────────────┐
│  INSTALAÇÃO     │           │  ROLLBACK AUTOMÁTICO│
│  CONCLUÍDA      │           │  ✓ Remove parcial   │
│                 │           │  ✓ Restaura backup  │
│                 │           │  ✓ Valida restauro  │
└─────────────────┘           └─────────────────────┘
```

---

## 🧪 Testes de Segurança

Execute os testes de segurança:

```powershell
# Executa todos os testes de segurança
.\tests\Security.Tests.ps1
```

**Cobertura de Testes**:
- ✅ Validação de tamanho de arquivos
- ✅ Validação de integridade
- ✅ Criação e validação de backups
- ✅ Rollback automático
- ✅ Validação de origem e destino
- ✅ Simulação de cenários de falha
- ✅ Validação de checksums

---

## 📝 Logs e Auditoria

Todas as operações são registradas em logs detalhados:

**chainsaw_installer.cmd**:
- Log salvo em: `chainsaw_installer_YYYYMMDD_HHMMSS.log`
- Copiado para: `%INSTALL_DIR%\installation\inst_docs\inst_logs\`

**install.ps1**:
- Log salvo em: `installation\inst_docs\inst_logs\install_YYYYMMDD_HHMMSS.log`

**Informações Registradas**:
- ✓ Timestamp de cada operação
- ✓ Validações executadas
- ✓ Tamanhos de arquivos
- ✓ Caminhos de backup
- ✓ Erros e avisos
- ✓ Operações de rollback

---

## ⚠️ Mensagens de Erro

### Erro Crítico: Backup Falhou
```
[ERRO CRITICO] Falha ao criar backup de seguranca!
[ERRO] NAO E SEGURO CONTINUAR sem backup valido.
[ERRO] Instalacao ABORTADA para proteger seus dados.
```
**Ação**: Feche programas que possam estar usando arquivos e tente novamente.

### Erro Crítico: Conteúdo Inválido
```
[ERRO CRITICO] Conteudo extraido INVALIDO ou INCOMPLETO!
[ERRO] NAO E SEGURO instalar arquivos incompletos.
[ERRO] Instalacao ABORTADA para proteger sua instalacao atual.
```
**Ação**: Verifique conexão de internet e tente novamente.

### Rollback Ativado
```
[ERRO CRITICO] Falha ao copiar arquivos para o destino!
[ROLLBACK] Tentando restaurar backup...
[OK] Backup restaurado com sucesso!
[INFO] Sistema retornou ao estado anterior.
```
**Ação**: Verifique logs para identificar causa da falha.

---

## 🛠️ Recuperação Manual

Se o rollback automático falhar, o backup está preservado:

**Localização do Backup**:
```
%USERPROFILE%\chainsaw_backup_YYYYMMDD_HHMMSS\
```

**Restauração Manual**:
1. Navegue até a pasta de backup
2. Copie todo o conteúdo
3. Cole em `%USERPROFILE%\chainsaw\`

---

## ✅ Garantias de Segurança

1. **NUNCA** remove dados sem backup validado
2. **NUNCA** instala conteúdo sem validação completa
3. **SEMPRE** valida origem antes de copiar
4. **SEMPRE** valida destino após copiar
5. **SEMPRE** mantém backup até confirmação de sucesso
6. **SEMPRE** executa rollback automático em caso de falha
7. **SEMPRE** registra todas as operações em log

---

## 📞 Suporte

Em caso de problemas:
1. Verifique o arquivo de log mais recente
2. Verifique se há backups em `%USERPROFILE%\CHAINSAW\backups\`
3. Reporte o problema com o conteúdo do log

---

**Última atualização**: 25 de novembro de 2025  
**Versão do documento**: 1.0
