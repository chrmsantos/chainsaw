# API Documentation

Esta documentação descreve as principais funções e procedimentos disponíveis no CHAINSAW PROPOSITURAS.

## Índice

- [Função Principal](#função-principal)
- [Funções de Configuração](#funções-de-configuração)
- [Funções de Backup](#funções-de-backup)
- [Funções de Formatação](#funções-de-formatação)
- [Funções de Validação](#funções-de-validação)
- [Funções de Logging](#funções-de-logging)
- [Constantes](#constantes)

## Função Principal

### `PadronizarDocumentoMain()`

**Descrição:** Função principal do sistema que executa todo o processo de padronização do documento.

**Sintaxe:**
```vba
Public Sub PadronizarDocumentoMain()
```

**Parâmetros:** Nenhum

**Retorna:** Nada (Sub)

**Exemplo de Uso:**
```vba
Call PadronizarDocumentoMain
```

**Funcionalidades:**
- Carrega configurações do arquivo INI
- Executa validações de segurança
- Cria backup automático
- Aplica formatação padronizada
- Executa limpeza de elementos visuais
- Gera logs do processo

## Funções de Configuração

### `LoadConfigFromFile(filePath As String) As Boolean`

**Descrição:** Carrega configurações de um arquivo INI específico.

**Sintaxe:**
```vba
Private Function LoadConfigFromFile(filePath As String) As Boolean
```

**Parâmetros:**
- `filePath` (String): Caminho completo para o arquivo de configuração

**Retorna:** 
- `Boolean`: True se as configurações foram carregadas com sucesso, False caso contrário

**Exemplo:**
```vba
Dim success As Boolean
success = LoadConfigFromFile("C:\Config\chainsaw-config.ini")
```

### `GetConfigValue(section As String, key As String, defaultValue As String) As String`

**Descrição:** Obtém um valor específico da configuração.

**Sintaxe:**
```vba
Private Function GetConfigValue(section As String, key As String, defaultValue As String) As String
```

**Parâmetros:**
- `section` (String): Seção do arquivo INI
- `key` (String): Chave da configuração
- `defaultValue` (String): Valor padrão se a chave não for encontrada

**Retorna:**
- `String`: Valor da configuração ou valor padrão

## Funções de Backup

### `CreateDocumentBackup() As String`

**Descrição:** Cria um backup do documento atual antes do processamento.

**Sintaxe:**
```vba
Private Function CreateDocumentBackup() As String
```

**Parâmetros:** Nenhum

**Retorna:**
- `String`: Caminho do arquivo de backup criado ou string vazia em caso de erro

**Exemplo:**
```vba
Dim backupPath As String
backupPath = CreateDocumentBackup()
If Len(backupPath) > 0 Then
    Debug.Print "Backup criado em: " & backupPath
End If
```

### `RestoreFromBackup(backupPath As String) As Boolean`

**Descrição:** Restaura o documento a partir de um backup específico.

**Sintaxe:**
```vba
Public Function RestoreFromBackup(backupPath As String) As Boolean
```

**Parâmetros:**
- `backupPath` (String): Caminho completo para o arquivo de backup

**Retorna:**
- `Boolean`: True se a restauração foi bem-sucedida, False caso contrário

## Funções de Formatação

### `ApplyStandardFormatting() As Boolean`

**Descrição:** Aplica a formatação padrão ao documento.

**Sintaxe:**
```vba
Private Function ApplyStandardFormatting() As Boolean
```

**Parâmetros:** Nenhum

**Retorna:**
- `Boolean`: True se a formatação foi aplicada com sucesso

**Funcionalidades:**
- Configura margens padrão
- Aplica fonte institucional
- Formata parágrafos especiais
- Adiciona cabeçalho e rodapé

### `FormatParagraphsByType() As Boolean`

**Descrição:** Formata parágrafos específicos baseado no tipo de proposição.

**Sintaxe:**
```vba
Private Function FormatParagraphsByType() As Boolean
```

**Parâmetros:** Nenhum

**Retorna:**
- `Boolean`: True se a formatação foi aplicada com sucesso

## Funções de Validação

### `ValidateDocumentIntegrity() As Boolean`

**Descrição:** Valida a integridade do documento antes do processamento.

**Sintaxe:**
```vba
Private Function ValidateDocumentIntegrity() As Boolean
```

**Parâmetros:** Nenhum

**Retorna:**
- `Boolean`: True se o documento passou em todas as validações

**Validações Realizadas:**
- Verifica se o documento não está protegido
- Confirma versão compatível do Word
- Valida espaço em disco disponível
- Verifica tipo de proposição

### `CheckWordVersion() As Boolean`

**Descrição:** Verifica se a versão do Word é compatível.

**Sintaxe:**
```vba
Private Function CheckWordVersion() As Boolean
```

**Parâmetros:** Nenhum

**Retorna:**
- `Boolean`: True se a versão é compatível (Word 2010 ou superior)

## Funções de Logging

### `LogMessage(message As String, level As String)`

**Descrição:** Registra uma mensagem no sistema de logging.

**Sintaxe:**
```vba
Private Sub LogMessage(message As String, level As String)
```

**Parâmetros:**
- `message` (String): Mensagem a ser registrada
- `level` (String): Nível do log (ERROR, WARNING, INFO, DEBUG)

**Exemplo:**
```vba
LogMessage "Processo iniciado", LOG_LEVEL_INFO
LogMessage "Erro na validação: " & Err.Description, LOG_LEVEL_ERROR
```

### `InitializeLogging() As Boolean`

**Descrição:** Inicializa o sistema de logging.

**Sintaxe:**
```vba
Private Function InitializeLogging() As Boolean
```

**Parâmetros:** Nenhum

**Retorna:**
- `Boolean`: True se o sistema de logging foi inicializado com sucesso

## Constantes

### Constantes de Sistema

```vba
Private Const VERSION As String = "v1.9.1-Alpha-8"
Private Const SYSTEM_NAME As String = "CHAINSAW PROPOSITURAS"
Private Const CONFIG_FILE_NAME As String = "\chainsaw-config.ini"
```

### Constantes de Log

```vba
Private Const LOG_LEVEL_ERROR As String = "ERROR"
Private Const LOG_LEVEL_WARNING As String = "WARNING"
Private Const LOG_LEVEL_INFO As String = "INFO"
Private Const LOG_LEVEL_DEBUG As String = "DEBUG"
```

### Constantes de Performance

```vba
Private Const MAX_PARAGRAPH_BATCH_SIZE As Long = 50
Private Const MAX_FIND_REPLACE_BATCH As Long = 100
Private Const OPTIMIZATION_THRESHOLD As Long = 1000
```

### Constantes de Erro

```vba
Private Const ERR_WORD_NOT_FOUND As Long = 5000
Private Const ERR_INCOMPATIBLE_VERSION As Long = 5001
Private Const ERR_DOCUMENT_PROTECTED As Long = 5002
Private Const ERR_BACKUP_FAILED As Long = 5003
Private Const ERR_INVALID_DOCUMENT As Long = 5004
```

## Estrutura de Configuração

### Exemplo de Arquivo de Configuração

```ini
[GERAL]
debug_mode = false
performance_mode = true
compatibility_mode = true

[VALIDACOES]
validate_document_integrity = true
validate_proposition_type = true
check_word_version = true
min_word_version = 14.0

[BACKUP]
auto_backup = true
backup_before_processing = true
max_backup_files = 10

[FORMATACAO]
apply_page_setup = true
apply_standard_font = true
apply_standard_paragraphs = true

[PERFORMANCE]
disable_screen_updating = true
use_bulk_operations = true
batch_paragraph_operations = true
```

## Tratamento de Erros

### Padrão de Tratamento

```vba
Public Function MinhaFuncao() As Boolean
    On Error GoTo ErrorHandler
    
    ' Código principal aqui
    
    MinhaFuncao = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro em MinhaFuncao: " & Err.Description, LOG_LEVEL_ERROR
    MinhaFuncao = False
End Function
```

### Códigos de Erro Customizados

| Código | Constante | Descrição |
|--------|-----------|-----------|
| 5000 | ERR_WORD_NOT_FOUND | Microsoft Word não encontrado |
| 5001 | ERR_INCOMPATIBLE_VERSION | Versão incompatível do Word |
| 5002 | ERR_DOCUMENT_PROTECTED | Documento protegido |
| 5003 | ERR_BACKUP_FAILED | Falha na criação de backup |
| 5004 | ERR_INVALID_DOCUMENT | Documento inválido |

## Performance e Otimizações

### Processamento em Lote

O sistema utiliza processamento em lote para melhorar a performance:

- **Parágrafos**: Processados em grupos de até 50 itens
- **Find/Replace**: Operações agrupadas em até 100 operações
- **Screen Updating**: Desabilitado durante processamento intensivo

### Gestão de Memória

- Coleta automática de lixo após operações intensivas
- Cache de objetos frequentemente acessados
- Minimização de criação/destruição de objetos

## Exemplos de Uso Avançado

### Execução com Configuração Customizada

```vba
Sub ExecutarComConfiguracaoCustomizada()
    ' Carregar configuração específica
    If LoadConfigFromFile("C:\Configs\config-especial.ini") Then
        Call PadronizarDocumentoMain
    Else
        MsgBox "Erro ao carregar configuração personalizada"
    End If
End Sub
```

### Processamento com Callback de Progresso

```vba
Sub ProcessarComProgresso()
    Dim i As Long
    Dim totalSteps As Long
    
    totalSteps = 10 ' Número de etapas
    
    For i = 1 To totalSteps
        ' Atualizar progresso
        LogMessage "Processando etapa " & i & " de " & totalSteps, LOG_LEVEL_INFO
        
        ' Executar etapa específica aqui
        
        ' Atualizar interface (se necessário)
        DoEvents
    Next i
End Sub
```

---

**Nota:** Esta documentação refere-se à versão 1.9.1-Alpha-8 do CHAINSAW PROPOSITURAS. Para versões mais recentes, consulte a documentação atualizada no repositório oficial.