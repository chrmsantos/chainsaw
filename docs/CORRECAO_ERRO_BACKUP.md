# Correção: Erro de Acesso Negado ao Criar Backup

## 🐛 Problema Identificado

No log `install_20251105_151951.log`, foi identificado o seguinte erro:

```
[2025-11-05 15:19:54] [ERROR] Erro ao criar backup: O acesso ao caminho 
'C:\Users\csantos\AppData\Roaming\Microsoft\Templates' foi negado.
```

### Causa Raiz

O erro ocorreu porque:

1. **Word pode estar aberto**: Arquivos na pasta Templates podem estar em uso pelo Microsoft Word
2. **Arquivos bloqueados**: Alguns arquivos podem estar bloqueados por outros processos
3. **Operação Rename-Item**: O método `Rename-Item` falha quando arquivos estão em uso

## ✅ Correção Implementada

### 1. Verificação do Word Antes do Backup

Adicionada verificação se o Word está em execução:

```powershell
if (Test-WordRunning) {
    # Avisa o usuário
    # Aguarda fechamento do Word
    # Verifica novamente antes de continuar
}
```

### 2. Método de Backup Alternativo

Implementado fallback quando `Rename-Item` falha:

```powershell
try {
    # Método 1: Rename-Item (mais rápido)
    Rename-Item -Path $SourceFolder -NewName $backupName -Force
}
catch [System.IO.IOException] {
    # Método 2: Copy + Delete (mais robusto)
    Copy-Item -Path $SourceFolder -Destination $backupPath -Recurse -Force
    Start-Sleep -Seconds 1  # Aguarda liberação de arquivos
    Remove-Item -Path $SourceFolder -Recurse -Force
}
```

### 3. Função Test-WordRunning Movida

A função `Test-WordRunning` foi movida para antes de `Backup-TemplatesFolder` para estar disponível quando necessária.

**Estrutura Atualizada:**
```
Funções Auxiliares
├── Test-WordRunning          ← Movida para cá
│
Funções de Backup
├── Backup-TemplatesFolder    ← Agora pode usar Test-WordRunning
└── Remove-OldBackups
│
Funções de Importação
├── Test-CustomizationsAvailable
├── Import-NormalTemplate
└── ...
```

## 🎯 Como Funciona Agora

### Fluxo de Backup Melhorado

```
1. Verificar se Word está aberto
   ├── Se SIM → Avisar usuário → Aguardar fechamento
   └── Se NÃO → Continuar

2. Tentar Rename-Item (método rápido)
   ├── Se SUCESSO → Backup criado ✓
   └── Se FALHA (arquivo em uso) → Ir para passo 3

3. Método alternativo: Copy + Delete
   ├── Copiar pasta inteira
   ├── Aguardar 1 segundo
   ├── Deletar pasta original
   └── Backup criado ✓
```

## 📋 Mensagens ao Usuário

Quando o Word está aberto, o usuário verá:

```
╔════════════════════════════════════════════════════════════════╗
║                  ⚠ MICROSOFT WORD ABERTO ⚠                    ║
╚════════════════════════════════════════════════════════════════╝

O Microsoft Word está em execução e deve ser fechado antes de
continuar com a instalação.

Por favor:
  1. Salve todos os documentos abertos no Word
  2. Feche completamente o Microsoft Word
  3. Pressione qualquer tecla para continuar
```

## 🧪 Testes Recomendados

### Cenário 1: Word Fechado
```cmd
# Certifique-se que o Word está fechado
cd %USERPROFILE%\Documents\chainsaw
install.cmd
```

**Resultado Esperado:** Backup criado com sucesso usando Rename-Item (rápido)

### Cenário 2: Word Aberto
```cmd
# Abra o Word antes de executar
cd %USERPROFILE%\Documents\chainsaw
install.cmd
```

**Resultado Esperado:** 
1. Script detecta Word aberto
2. Exibe aviso
3. Aguarda fechamento
4. Continua após Word ser fechado

### Cenário 3: Arquivo em Uso (Sem Word)
```cmd
# Se algum arquivo estiver em uso por outro processo
cd %USERPROFILE%\Documents\chainsaw
install.cmd
```

**Resultado Esperado:** 
1. Rename-Item falha
2. Método alternativo (Copy + Delete) é usado
3. Backup criado com sucesso

## 🔍 Verificação de Logs

Após executar, verifique o log em:
```
%USERPROFILE%\chainsaw\logs\install_[timestamp].log
```

### Log de Sucesso - Método Rápido

```log
[INFO] Criando backup da pasta Templates...
[INFO] Origem: C:\Users\csantos\AppData\Roaming\Microsoft\Templates
[INFO] Destino: C:\Users\csantos\AppData\Roaming\Microsoft\Templates_backup_20251105_152500
[SUCCESS] Backup criado com sucesso: Templates_backup_20251105_152500 ✓
```

### Log de Sucesso - Método Alternativo

```log
[INFO] Criando backup da pasta Templates...
[WARNING] Erro de acesso ao renomear (possível arquivo em uso)
[INFO] Tentando método alternativo (cópia)...
[SUCCESS] Backup criado com sucesso (método cópia): Templates_backup_20251105_152500 ✓
```

### Log com Word Aberto

```log
[WARNING] Aguardando fechamento do Word...
[INFO] Word fechado, continuando...
[INFO] Criando backup da pasta Templates...
[SUCCESS] Backup criado com sucesso: Templates_backup_20251105_152500 ✓
```

## 💡 Dicas para Evitar o Erro

### Antes de Executar install.cmd

1. ✅ **Feche o Microsoft Word completamente**
   - Salve todos os documentos
   - Feche todas as janelas do Word
   - Verifique no Gerenciador de Tarefas se `WINWORD.EXE` não está em execução

2. ✅ **Feche outros aplicativos do Office**
   - Outlook (se usa modelos do Word)
   - PowerPoint (se compartilha recursos)
   - Excel (se usa templates do Word)

3. ✅ **Execute como usuário normal**
   - NÃO use "Executar como administrador"
   - Use sua sessão de usuário normal

### Durante a Instalação

- ⏳ Se solicitado, aguarde o script completar
- 🚫 Não abra o Word durante a instalação
- 📝 Acompanhe as mensagens na tela

## 🆘 Troubleshooting

### Erro Persiste Mesmo com Word Fechado

**Solução:**

1. Abra o Gerenciador de Tarefas (Ctrl + Shift + Esc)
2. Vá para aba "Detalhes"
3. Procure por `WINWORD.EXE`
4. Se encontrar, clique com botão direito → "Finalizar tarefa"
5. Execute install.cmd novamente

### Erro "O acesso ao caminho foi negado" Continua

**Possíveis causas:**

1. **Antivírus bloqueando**: Temporariamente desabilite o antivírus
2. **Sincronização de nuvem**: OneDrive/Google Drive podem bloquear arquivos
3. **Permissões**: Verifique se tem permissão de escrita em `%APPDATA%`

**Solução alternativa:**

```powershell
# Verificar permissões
$templatesPath = "$env:APPDATA\Microsoft\Templates"
$acl = Get-Acl $templatesPath
$acl.Access | Format-Table IdentityReference, FileSystemRights

# Se necessário, tomar propriedade
takeown /f $templatesPath /r /d y
icacls $templatesPath /grant "$env:USERNAME:(OI)(CI)F" /t
```

## 📊 Mudanças no Código

### Arquivos Modificados

- ✅ `install.ps1` - Versão 2.0.0
  - Função `Backup-TemplatesFolder` melhorada
  - Função `Test-WordRunning` movida
  - Método de backup alternativo adicionado
  - Verificação de Word em execução adicionada

### Linhas Modificadas

| Função | Linhas Adicionadas | Impacto |
|--------|-------------------|---------|
| `Test-WordRunning` | ~6 | Movida para antes de Backup |
| `Backup-TemplatesFolder` | ~50 | Verificação de Word + método alternativo |

## ✅ Status

- [x] Erro identificado
- [x] Causa raiz determinada
- [x] Correção implementada
- [x] Sintaxe validada
- [x] Documentação criada
- [ ] Teste em ambiente real (próximo passo)

## 🚀 Próximo Passo

Execute a instalação novamente:

```cmd
cd %USERPROFILE%\Documents\chainsaw
install.cmd
```

Se o erro persistir, verifique:
1. Word está fechado?
2. Gerenciador de Tarefas mostra WINWORD.EXE?
3. Antivírus está bloqueando?
4. Tem permissões na pasta Templates?

---

**Correção aplicada em:** 05/11/2025  
**Versão do script:** 2.0.0  
**Status:** ✅ Pronto para teste
