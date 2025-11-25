# ✅ Checklist de Verificação - Proteções de Segurança

Use este checklist para validar que todas as proteções estão funcionando corretamente.

---

## 📋 PRÉ-INSTALAÇÃO

### Validação de Código
- [ ] Executar testes de segurança: `.\tests\Security.Tests.ps1`
- [ ] Verificar que não há erros de compilação
- [ ] Revisar logs de testes anteriores

### Preparação do Ambiente
- [ ] Fechar todas as instâncias do Microsoft Word
- [ ] Verificar espaço em disco disponível (mínimo 500MB)
- [ ] Verificar permissões de escrita no perfil do usuário

---

## 🧪 TESTES DE VALIDAÇÃO

### Teste 1: Download Corrompido
```powershell
# Simular download corrompido (arquivo muito pequeno)
# Resultado esperado: Instalação ABORTADA com mensagem clara
```
- [ ] Criar arquivo ZIP de teste < 100KB
- [ ] Executar installer
- [ ] Verificar que instalação foi abortada
- [ ] Verificar mensagem: "Arquivo ZIP muito pequeno"

### Teste 2: Conteúdo Incompleto
```powershell
# Simular conteúdo extraído sem arquivos essenciais
# Resultado esperado: Instalação ABORTADA após validação
```
- [ ] Criar ZIP sem pasta `installation/inst_scripts/`
- [ ] Executar installer
- [ ] Verificar que instalação foi abortada
- [ ] Verificar mensagem: "Conteúdo extraído INVÁLIDO"

### Teste 3: Backup Obrigatório
```powershell
# Verificar criação e validação de backup
# Resultado esperado: Backup criado e validado antes de modificações
```
- [ ] Executar instalação normal
- [ ] Verificar mensagem: "Criando backup OBRIGATÓRIO"
- [ ] Verificar pasta de backup criada
- [ ] Verificar mensagem: "Backup validado: X arquivos"

### Teste 4: Rollback Automático
```powershell
# Simular falha durante cópia
# Resultado esperado: Backup restaurado automaticamente
```
- [ ] Criar cenário de falha (ex: disco cheio simulado)
- [ ] Verificar mensagem: "[ROLLBACK] Tentando restaurar backup"
- [ ] Verificar que backup foi restaurado
- [ ] Verificar que sistema voltou ao estado anterior

---

## 🔍 VALIDAÇÕES DO CÓDIGO

### chainsaw_installer.cmd

#### Validação de Download
- [ ] Linha ~106-142: Validação de tamanho do ZIP
  ```batch
  if %ZIP_SIZE% LSS 102400 (
  ```

- [ ] Linha ~143-199: Teste de integridade do ZIP
  ```batch
  System.IO.Compression.ZipFile
  ```

- [ ] Linha ~200-222: Validação de conteúdo mínimo
  ```batch
  if $entryCount -lt 10
  ```

#### Backup Obrigatório
- [ ] Linha ~147-156: Mensagem "Backup OBRIGATORIO"
- [ ] Linha ~159-200: Criação de backup com fallback
- [ ] Linha ~202-222: Validação do backup
  ```batch
  BACKUP_FILE_COUNT
  if %BACKUP_FILE_COUNT% LSS 5
  ```

#### Validação de Conteúdo
- [ ] Linha ~224-280: Validação de estrutura
  ```batch
  if not exist "%SOURCE_DIR%\installation"
  if not exist "%SOURCE_DIR%\installation\inst_scripts\install.cmd"
  ```

- [ ] Linha ~250-265: Contagem de arquivos
  ```batch
  EXTRACTED_FILE_COUNT
  if %EXTRACTED_FILE_COUNT% LSS 20
  ```

#### Rollback
- [ ] Linha ~314-338: Detecção de falha e rollback
  ```batch
  [ROLLBACK] Tentando restaurar backup
  ```

- [ ] Linha ~340-365: Validação final
  ```batch
  FINAL_VALIDATION_FAILED
  ```

### install.ps1

#### Copy-StampFile
- [ ] Validação 1: Arquivo existe
  ```powershell
  if (-not (Test-Path $SourceFile))
  ```

- [ ] Validação 2: Tamanho mínimo
  ```powershell
  if ($sourceFileInfo.Length -lt 100)
  ```

- [ ] Validação 3: Cópia bem-sucedida
  ```powershell
  if ($sourceSize -ne $destSize)
  ```

#### Copy-TemplatesFolder
- [ ] Validação 1: Pasta existe
  ```powershell
  if (-not (Test-Path $SourceFolder))
  ```

- [ ] Validação 2: Pasta não vazia
  ```powershell
  if ($sourceItems.Count -eq 0)
  ```

- [ ] Validação 3: Normal.dotm presente
  ```powershell
  if (-not (Test-Path $sourceNormalDotm))
  ```

- [ ] Validação 4: Tamanho de Normal.dotm
  ```powershell
  if ($normalDotmSize -lt 10000)
  ```

- [ ] Validação 5: Cópia de Normal.dotm validada
  ```powershell
  if ($destNormalDotmSize -ne $normalDotmSize)
  ```

#### Rollback
- [ ] Validação de backup antes de restaurar
  ```powershell
  if ($backupItems.Count -eq 0)
  ```

- [ ] Restauração validada
  ```powershell
  if (Test-Path $restoredPath)
  ```

---

## 📊 VERIFICAÇÃO DE LOGS

### Log do Installer (chainsaw_installer.cmd)
- [ ] Verificar criação de log: `chainsaw_installer_*.log`
- [ ] Verificar presença de timestamps
- [ ] Verificar registro de todas as etapas
- [ ] Verificar mensagens de validação

**Etapas esperadas**:
1. Download do código-fonte
2. Validação do arquivo baixado
3. Backup obrigatório da instalação existente
4. Validação do backup
5. Extração e validação dos arquivos
6. Instalação
7. Validação final

### Log do install.ps1
- [ ] Verificar criação de log em `installation\inst_docs\inst_logs\`
- [ ] Verificar registro de operações de cópia
- [ ] Verificar mensagens de validação
- [ ] Verificar registro de sucessos e erros

**Operações esperadas**:
- Cópia de stamp.png (com validações)
- Backup de Templates
- Cópia de Templates (com validações)
- Importação de módulo VBA

---

## 🎯 CENÁRIOS DE TESTE COMPLETOS

### Cenário 1: Instalação Limpa (Sem instalação anterior)
```
✓ Download validado
✓ Conteúdo extraído validado
✓ Nenhuma pasta anterior (sem backup necessário)
✓ Instalação concluída com sucesso
```
- [ ] Executado
- [ ] Sucesso confirmado
- [ ] Logs verificados

### Cenário 2: Atualização (Com instalação anterior)
```
✓ Download validado
✓ Conteúdo extraído validado
✓ Backup criado e validado
✓ Instalação concluída com sucesso
✓ Backup preservado
```
- [ ] Executado
- [ ] Backup criado
- [ ] Sucesso confirmado
- [ ] Backup preservado

### Cenário 3: Falha no Download
```
✓ Download falha (arquivo corrompido)
✓ Validação detecta problema
✓ Instalação ABORTADA
✓ Nenhuma modificação feita
```
- [ ] Simulado
- [ ] Instalação abortada
- [ ] Mensagem clara exibida
- [ ] Sistema inalterado

### Cenário 4: Falha na Instalação
```
✓ Download OK
✓ Validação OK
✓ Backup criado e validado
✓ Falha durante cópia
✓ Rollback automático ativado
✓ Backup restaurado
✓ Sistema retornou ao estado anterior
```
- [ ] Simulado
- [ ] Rollback ativado
- [ ] Backup restaurado
- [ ] Sistema recuperado

---

## 📝 VERIFICAÇÃO DE DOCUMENTAÇÃO

### Arquivos de Documentação
- [ ] `docs/PROTECOES_SEGURANCA.md` existe e está completo
- [ ] `docs/CHANGELOG_SEGURANCA.md` existe e está completo
- [ ] `RESUMO_SEGURANCA.md` existe e está completo
- [ ] Este checklist está completo

### Conteúdo da Documentação
- [ ] Descrição do problema está clara
- [ ] Soluções implementadas estão documentadas
- [ ] Fluxo de segurança está visual
- [ ] Cenários protegidos estão listados
- [ ] Mensagens de erro estão documentadas
- [ ] Procedimentos de recuperação estão claros

---

## 🧪 TESTES AUTOMATIZADOS

### Arquivo de Testes
- [ ] `tests/Security.Tests.ps1` existe
- [ ] Todos os testes passam sem erros
- [ ] Cobertura de código adequada

### Executar Testes
```powershell
cd c:\Users\csantos\chainsaw
.\tests\Security.Tests.ps1
```

**Resultado esperado**:
```
✅ Todos os testes passaram
✅ Nenhum erro encontrado
✅ Proteções validadas
```

- [ ] Testes executados
- [ ] Todos os testes passaram
- [ ] Relatório de testes salvo

---

## ✅ VALIDAÇÃO FINAL

### Checklist de Aprovação
- [ ] Todos os testes de segurança passaram
- [ ] Documentação completa e revisada
- [ ] Cenários críticos testados
- [ ] Logs verificados
- [ ] Backups funcionando
- [ ] Rollback testado
- [ ] Validações confirmadas

### Assinatura
- [ ] Desenvolvedor: _________________________
- [ ] Data: ___/___/_____
- [ ] Versão testada: 2.0.3

---

## 📞 PRÓXIMOS PASSOS

Após completar este checklist:

1. [ ] Commit das alterações no repositório
2. [ ] Tag de versão: `v2.0.3`
3. [ ] Atualizar CHANGELOG principal
4. [ ] Comunicar mudanças à equipe
5. [ ] Monitorar primeira instalação em produção

---

## 🚨 EM CASO DE FALHA

Se qualquer item deste checklist falhar:

1. **NÃO** prosseguir para produção
2. Investigar a falha nos logs
3. Corrigir o problema
4. Reiniciar o checklist
5. Documentar o problema e solução

---

**Data de criação**: 25 de novembro de 2025  
**Versão do checklist**: 1.0  
**Status**: ✅ Pronto para uso
