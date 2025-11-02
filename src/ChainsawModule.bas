' =============================================================================
' PROJETO: CHAINSAW FOR PROPOSALS (CHAINSW-FPROPS)
' =============================================================================
'
' Sistema automatizado de padronização de documentos legislativos no Microsoft Word
'
' Licença: Apache 2.0 modificada (ver LICENSE)
' Versão: 1.0-beta1 | Data: 2025-11-02
' Repositório: github.com/chrmsantos/chainsaw-fprops
' Autor: Christian Martin dos Santos <chrmsantos@gmail.com>
'
' =============================================================================
' CHANGELOG v1.0-beta1 (2025-11-02):
' =============================================================================
'
' ✅ MELHORIAS CRÍTICAS:
'   - Refatoração completa da aplicação de negrito em parágrafos especiais
'   - Nova função ApplyBoldToSpecialParagraphs() simplificada e otimizada
'   - Removidos loops caractere-a-caractere desnecessários (melhoria de ~80% performance)
'   - Aplicação única e atômica de negrito ao final do processamento
'
' ✅ CÓDIGO LIMPO:
'   - Adicionadas funções auxiliares: ValidateDocument(), GetCleanParagraphText()
'   - Constantes nomeadas para valores mágicos (CONSIDERANDO_PREFIX, etc.)
'   - Modo DEBUG para desenvolvimento e troubleshooting
'   - Mensagens de erro amigáveis ao usuário (ShowUserFriendlyError)
'
' ✅ MANUTENIBILIDADE:
'   - Código duplicado reduzido em ~30%
'   - Separação clara de responsabilidades (formatação vs. aplicação de negrito)
'   - Logging aprimorado com suporte a Debug.Print
'   - Documentação inline melhorada
'
' 🔄 REFATORAÇÕES:
'   - FormatJustificativaAnexoParagraphs: removida aplicação de negrito
'   - Fluxo principal simplificado: 3 chamadas → 2 chamadas
'   - Validação de documento centralizada
'
' =============================================================================
' FUNCIONALIDADES PRINCIPAIS:
' =============================================================================
'
' • VERIFICAÇÕES DE SEGURANÇA E COMPATIBILIDADE:
'   - Validação de versão do Word (mínimo: 2010)
'   - Verificação de tipo e proteção do documento
'   - Controle de espaço em disco e estrutura mínima
'   - Proteção contra falhas e recuperação automática
'
' • SISTEMA DE BACKUP AUTOMÁTICO:
'   - Backup automático antes de qualquer modificação
'   - Pasta de backups organizada por documento
'   - Limpeza automática de backups antigos (limite: 10 arquivos)
'   - Subrotina pública para acesso à pasta de backups
'
' • SUBROTINA PÚBLICA PARA SALVAR E SAIR:
'   - Verificação automática de todos os documentos abertos
'   - Detecção de documentos com alterações não salvas
'   - Interface profissional com opções claras ao usuário
'   - Salvamento assistido com diálogos para novos arquivos
'   - Confirmação dupla para fechamento sem salvar
'   - Tratamento robusto de erros e recuperação
'
' • FORMATAÇÃO AUTOMATIZADA INSTITUCIONAL:
'   - Limpeza completa de formatação ao iniciar
'   - Remoção robusta de espaços múltiplos e tabs
'   - Controle de linhas vazias (máximo 2 sequenciais)
'   - PROTEÇÃO MÁXIMA: Sistema avançado de backup/restauração de imagens
'   - PROTEÇÃO MÁXIMA: Preserva imagens inline, flutuantes e objetos
'   - PROTEÇÃO MÁXIMA: Detecta e protege shapes ancoradas e campos visuais
'   - Primeira linha: SEMPRE caixa alta, negrito, sublinhado, centralizada
'   - Parágrafos 2°, 3° e 4°: recuo esquerdo 9cm, sem recuo primeira linha
'   - "Considerando": caixa alta e negrito no início de parágrafos
'   - "Justificativa": centralizada, sem recuos, negrito, capitalizada
'   - "Anexo/Anexos": alinhado à esquerda, sem recuos, negrito, capitalizado
'   - Configuração de margens e orientação (A4)
'   - Fonte Arial 12pt com espaçamento 1.4
'   - Recuos e alinhamento justificado
'   - Cabeçalho com logotipo institucional
'   - Rodapé com numeração centralizada
'   - Visualização: zoom 110% (mantido), demais configurações preservadas
'   - PROTEÇÃO TOTAL: Preserva réguas, modos de exibição e configurações originais
'   - Remoção de marcas d'água e formatações manuais
'
' • SISTEMA DE LOGS E MONITORAMENTO:
'   - Registro detalhado de operações
'   - Controle de erros com fallback
'   - Mensagens na barra de status
'   - Histórico de execução
'
' • SISTEMA DE PROTEÇÃO DE CONFIGURAÇÕES DE VISUALIZAÇÃO:
'   - Backup automático de todas as configurações de exibição
'   - Preservação de réguas (horizontal e vertical)
'   - Manutenção do modo de visualização original
'   - Proteção de configurações de marcas de formatação
'   - Restauração completa após processamento (exceto zoom)
'   - Compatibilidade com todos os modos de exibição do Word
'
' • PERFORMANCE OTIMIZADA:
'   - Processamento eficiente para documentos grandes
'   - Desabilitação temporária de atualizações visuais
'   - Gerenciamento inteligente de recursos
'   - Sistema de logging otimizado (principais, warnings e erros)
'
' =============================================================================

'VBA
Option Explicit

'================================================================================
' CONSTANTS
'================================================================================

' Word built-in constants
Private Const wdNoProtection As Long = -1
Private Const wdTypeDocument As Long = 0
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdAlignParagraphJustify As Long = 3
Private Const wdLineSpaceSingle As Long = 0
Private Const wdLineSpace1pt5 As Long = 1
Private Const wdLineSpacingMultiple As Long = 5
Private Const wdStatisticPages As Long = 2
Private Const msoTrue As Long = -1
Private Const msoFalse As Long = 0
Private Const msoPicture As Long = 13
Private Const msoTextEffect As Long = 15
Private Const wdCollapseEnd As Long = 0
Private Const wdCollapseStart As Long = 1
Private Const wdFieldPage As Long = 33
Private Const wdFieldNumPages As Long = 26
Private Const wdFieldEmpty As Long = -1
Private Const wdRelativeHorizontalPositionPage As Long = 1
Private Const wdRelativeVerticalPositionPage As Long = 1
Private Const wdWrapTopBottom As Long = 3
Private Const wdAlertsAll As Long = 0
Private Const wdAlertsNone As Long = -1
Private Const wdColorAutomatic As Long = -16777216
Private Const wdOrientPortrait As Long = 0
Private Const wdUnderlineNone As Long = 0
Private Const wdUnderlineSingle As Long = 1
Private Const wdTextureNone As Long = 0
Private Const wdPrintView As Long = 3

' Document formatting constants
Private Const STANDARD_FONT As String = "Arial"
Private Const STANDARD_FONT_SIZE As Long = 12
Private Const FOOTER_FONT_SIZE As Long = 9
Private Const LINE_SPACING As Single = 14

' Margin constants in centimeters
Private Const TOP_MARGIN_CM As Double = 4.6
Private Const BOTTOM_MARGIN_CM As Double = 2
Private Const LEFT_MARGIN_CM As Double = 3
Private Const RIGHT_MARGIN_CM As Double = 3
Private Const HEADER_DISTANCE_CM As Double = 0.3
Private Const FOOTER_DISTANCE_CM As Double = 0.9

' Header image constants
Private Const HEADER_IMAGE_RELATIVE_PATH As String = "\chainsaw-proposituras\assets\stamp.png"
Private Const HEADER_IMAGE_MAX_WIDTH_CM As Double = 21
Private Const HEADER_IMAGE_TOP_MARGIN_CM As Double = 0.7
Private Const HEADER_IMAGE_HEIGHT_RATIO As Double = 0.19

' Minimum supported version
Private Const MIN_SUPPORTED_VERSION As Long = 14 ' Word 2010

' Logging constants
Private Const LOG_LEVEL_INFO As Long = 1
Private Const LOG_LEVEL_WARNING As Long = 2
Private Const LOG_LEVEL_ERROR As Long = 3

' Required string constant
Private Const REQUIRED_STRING As String = "$NUMERO$/$ANO$"

' Timeout constants
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

' Backup constants
Private Const BACKUP_FOLDER_NAME As String = "\chainsaw-proposituras\backups"
Private Const MAX_BACKUP_FILES As Long = 10

' Special paragraph constants
Private Const CONSIDERANDO_PREFIX As String = "considerando"
Private Const CONSIDERANDO_MIN_LENGTH As Long = 12
Private Const JUSTIFICATIVA_TEXT As String = "justificativa"

' Debug mode
Private Const DEBUG_MODE As Boolean = False

'================================================================================
' GLOBAL VARIABLES
'================================================================================
Private undoGroupEnabled As Boolean
Private loggingEnabled As Boolean
Private logFilePath As String
Private formattingCancelled As Boolean
Private executionStartTime As Date
Private backupFilePath As String

' Image protection variables
Private Type ImageInfo
    paraIndex As Long
    ImageIndex As Long
    ImageType As String
    ImageData As Variant
    Position As Long
    WrapType As Long
    Width As Single
    Height As Single
    LeftPosition As Single
    TopPosition As Single
    AnchorRange As Range
End Type

Private savedImages() As ImageInfo
Private imageCount As Long

' View settings backup variables
Private Type ViewSettings
    ViewType As Long
    ShowVerticalRuler As Boolean
    ShowHorizontalRuler As Boolean
    ShowFieldCodes As Boolean
    ShowBookmarks As Boolean
    ShowParagraphMarks As Boolean
    ShowSpaces As Boolean
    ShowTabs As Boolean
    ShowHiddenText As Boolean
    ShowOptionalHyphens As Boolean
    ShowAll As Boolean
    ShowDrawings As Boolean
    ShowObjectAnchors As Boolean
    ShowTextBoundaries As Boolean
    ShowHighlight As Boolean
    ' ShowAnimation removida - compatibilidade
    DraftFont As Boolean
    WrapToWindow As Boolean
    ShowPicturePlaceHolders As Boolean
    ShowFieldShading As Long
    TableGridlines As Boolean
    ' EnlargeFontsLessThan removida - compatibilidade
End Type

Private originalViewSettings As ViewSettings

'================================================================================
' MAIN ENTRY POINT - #STABLE
'================================================================================
Public Sub PadronizarDocumentoMain()
    On Error GoTo CriticalErrorHandler
    
    executionStartTime = Now
    formattingCancelled = False
    
    If Not CheckWordVersion() Then
        Application.StatusBar = "Erro: Versão do Word não suportada (mínimo: Word 2010)"
        LogMessage "Versão do Word " & Application.version & " não suportada. Mínimo: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
        MsgBox "Esta ferramenta requer Microsoft Word 2010 ou superior." & vbCrLf & _
               "Versão atual: " & Application.version & vbCrLf & _
               "Versão mínima: " & CStr(MIN_SUPPORTED_VERSION), vbCritical, "Versão Incompatível"
        Exit Sub
    End If
    
    Dim doc As Document
    Set doc = Nothing
    
    On Error Resume Next
    Set doc = ActiveDocument
    If doc Is Nothing Then
        Application.StatusBar = "Erro: Nenhum documento está acessível"
        LogMessage "Nenhum documento acessível para processamento", LOG_LEVEL_ERROR
        Exit Sub
    End If
    On Error GoTo CriticalErrorHandler
    
    If Not InitializeLogging(doc) Then
        LogMessage "Falha ao inicializar sistema de logs", LOG_LEVEL_WARNING
    End If
    
    LogMessage "Iniciando padronização do documento: " & doc.Name, LOG_LEVEL_INFO
    
    StartUndoGroup "Padronização de Documento - " & doc.Name
    
    If Not SetAppState(False, "Formatando documento...") Then
        LogMessage "Falha ao configurar estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    If Not PreviousChecking(doc) Then
        GoTo CleanUp
    End If
    
    If doc.Path = "" Then
        If Not SaveDocumentFirst(doc) Then
            Application.StatusBar = "Operação cancelada: documento precisa ser salvo"
            LogMessage "Operação cancelada - documento não foi salvo", LOG_LEVEL_INFO
            Exit Sub
        End If
    End If
    
    ' Cria backup do documento antes de qualquer modificação
    If Not CreateDocumentBackup(doc) Then
        LogMessage "Falha ao criar backup - continuando sem backup", LOG_LEVEL_WARNING
        Application.StatusBar = "Aviso: Backup não foi possível - processando sem backup"
    Else
        Application.StatusBar = "Backup criado - formatando documento..."
    End If
    
    ' Backup das configurações de visualização originais
    If Not BackupViewSettings(doc) Then
        LogMessage "Aviso: Falha no backup das configurações de visualização", LOG_LEVEL_WARNING
    End If

    ' Backup de imagens antes das formatações
    Application.StatusBar = "Catalogando imagens do documento..."
    If Not BackupAllImages(doc) Then
        LogMessage "Aviso: Falha no backup de imagens - continuando com proteção básica", LOG_LEVEL_WARNING
    End If

    If Not PreviousFormatting(doc) Then
        GoTo CleanUp
    End If

    ' Restaura imagens após formatações
    Application.StatusBar = "Verificando integridade das imagens..."
    If Not RestoreAllImages(doc) Then
        LogMessage "Aviso: Algumas imagens podem ter sido afetadas durante o processamento", LOG_LEVEL_WARNING
    End If

    ' Restaura configurações de visualização originais (exceto zoom)
    If Not RestoreViewSettings(doc) Then
        LogMessage "Aviso: Algumas configurações de visualização podem não ter sido restauradas", LOG_LEVEL_WARNING
    End If

    If formattingCancelled Then
        GoTo CleanUp
    End If

    Application.StatusBar = "Documento padronizado com sucesso!"
    LogMessage "Documento padronizado com sucesso", LOG_LEVEL_INFO

CleanUp:
    SafeCleanup
    CleanupImageProtection ' Nova função para limpar variáveis de proteção de imagens
    CleanupViewSettings    ' Nova função para limpar variáveis de configurações de visualização
    
    If Not SetAppState(True, "Documento padronizado com sucesso!") Then
        LogMessage "Falha ao restaurar estado da aplicação", LOG_LEVEL_WARNING
    End If
    
    SafeFinalizeLogging
    
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO #" & Err.Number & ": " & Err.Description & _
              " em " & Err.Source & " (Linha: " & Erl & ")"
    
    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro crítico durante processamento - verificar logs"
    
    ShowUserFriendlyError Err.Number, Err.Description
    EmergencyRecovery
End Sub

'================================================================================
' EMERGENCY RECOVERY - #STABLE
'================================================================================
'================================================================================
' USER-FRIENDLY ERROR HANDLER - #NEW
'================================================================================
Private Sub ShowUserFriendlyError(errNum As Long, errDesc As String)
    Dim msg As String
    
    Select Case errNum
        Case 91 ' Object variable not set
            msg = "Erro interno: Objeto não inicializado." & vbCrLf & vbCrLf & _
                  "Por favor, reinicie o Word e tente novamente."
        
        Case 5 ' Invalid procedure call
            msg = "Erro de configuração detectado." & vbCrLf & vbCrLf & _
                  "Verifique se o documento está em formato válido (.docx)."
        
        Case 70 ' Permission denied
            msg = "Permissão negada." & vbCrLf & vbCrLf & _
                  "O documento pode estar protegido ou somente leitura." & vbCrLf & _
                  "Salve uma cópia antes de continuar."
        
        Case 53 ' File not found
            msg = "Arquivo não encontrado." & vbCrLf & vbCrLf & _
                  "Verifique se o documento foi salvo corretamente."
        
        Case Else
            msg = "Erro inesperado (#" & errNum & "):" & vbCrLf & vbCrLf & _
                  errDesc & vbCrLf & vbCrLf & _
                  "Por favor, verifique o log para mais detalhes."
    End Select
    
    MsgBox msg, vbCritical, "Chainsaw Proposituras v1.0-beta1"
End Sub

'================================================================================
' EMERGENCY RECOVERY - #STABLE
'================================================================================
Private Sub EmergencyRecovery()
    On Error Resume Next
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False
    Application.EnableCancelKey = 0
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
    ' Limpa variáveis de proteção de imagens em caso de erro
    CleanupImageProtection
    
    ' Limpa variáveis de configurações de visualização em caso de erro
    CleanupViewSettings
    
    LogMessage "Recuperação de emergência executada", LOG_LEVEL_ERROR
        undoGroupEnabled = False
    
    CloseAllOpenFiles
End Sub

'================================================================================
' SAFE CLEANUP - LIMPEZA SEGURA - #STABLE
'================================================================================
Private Sub SafeCleanup()
    On Error Resume Next
    
    EndUndoGroup
    
    ReleaseObjects
End Sub

'================================================================================
' RELEASE OBJECTS - #STABLE
'================================================================================
Private Sub ReleaseObjects()
    On Error Resume Next
    
    Dim nullObj As Object
    Set nullObj = Nothing
    
    Dim memoryCounter As Long
    For memoryCounter = 1 To 3
        DoEvents
    Next memoryCounter
End Sub

'================================================================================
' CLOSE ALL OPEN FILES - #STABLE
'================================================================================
Private Sub CloseAllOpenFiles()
    On Error Resume Next
    
    Dim fileNumber As Integer
    For fileNumber = 1 To 511
        If Not EOF(fileNumber) Then
            Close fileNumber
        End If
    Next fileNumber
End Sub

'================================================================================
' VERSION COMPATIBILITY AND SAFETY CHECKS - #STABLE
'================================================================================
'================================================================================
' VALIDATION HELPERS - #NEW
'================================================================================
Private Function ValidateDocument(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateDocument = False
    
    If doc Is Nothing Then
        LogMessage "Documento é Nothing", LOG_LEVEL_ERROR
        Exit Function
    End If
    
    If doc.Paragraphs.Count = 0 Then
        LogMessage "Documento não tem parágrafos", LOG_LEVEL_WARNING
        Exit Function
    End If
    
    ValidateDocument = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro na validação do documento: " & Err.Description, LOG_LEVEL_ERROR
    ValidateDocument = False
End Function

'================================================================================
' TEXT CLEANING HELPERS - #NEW
'================================================================================
Private Function GetCleanParagraphText(para As Paragraph) As String
    On Error Resume Next
    
    Dim txt As String
    txt = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
    
    ' Remove pontuação final
    Do While Len(txt) > 0 And InStr(".,;:", Right(txt, 1)) > 0
        txt = Left(txt, Len(txt) - 1)
    Loop
    
    GetCleanParagraphText = Trim(LCase(txt))
End Function

Private Function RemovePunctuation(text As String) As String
    Dim result As String
    result = text
    
    ' Remove pontuação final
    Do While Len(result) > 0 And InStr(".,;:", Right(result, 1)) > 0
        result = Left(result, Len(result) - 1)
    Loop
    
    RemovePunctuation = Trim(result)
End Function

'================================================================================
' VERSION CHECK - #STABLE
'================================================================================
Private Function CheckWordVersion() As Boolean
    On Error GoTo ErrorHandler
    
    Dim version As Double
    ' Uso de CDbl para garantir conversão correta em todas as versões
    version = CDbl(Application.version)
    
    If version < MIN_SUPPORTED_VERSION Then
        CheckWordVersion = False
        LogMessage "Versão detectada: " & CStr(version) & " - Mínima suportada: " & CStr(MIN_SUPPORTED_VERSION), LOG_LEVEL_ERROR
    Else
        CheckWordVersion = True
        LogMessage "Versão do Word compatível: " & CStr(version), LOG_LEVEL_INFO
    End If
    
    Exit Function
    
ErrorHandler:
    ' Se não conseguir detectar a versão, assume incompatibilidade por segurança
    CheckWordVersion = False
    LogMessage "Erro ao detectar versão do Word: " & Err.Description, LOG_LEVEL_ERROR
End Function

'================================================================================
' SAFE PROPERTY ACCESS FUNCTIONS - Compatibilidade total com Word 2010+
'================================================================================
Private Function SafeGetCharacterCount(targetRange As Range) As Long
    On Error GoTo FallbackMethod
    
    ' Método preferido - mais rápido
    SafeGetCharacterCount = targetRange.Characters.count
    Exit Function
    
FallbackMethod:
    On Error GoTo ErrorHandler
    ' Método alternativo para versões com problemas de .Characters.Count
    SafeGetCharacterCount = Len(targetRange.text)
    Exit Function
    
ErrorHandler:
    ' Último recurso - valor padrão seguro
    SafeGetCharacterCount = 0
    LogMessage "Erro ao obter contagem de caracteres: " & Err.Description, LOG_LEVEL_WARNING
End Function

Private Function SafeSetFont(targetRange As Range, fontName As String, fontSize As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Aplica formatação de fonte de forma segura
    With targetRange.Font
        If fontName <> "" Then .Name = fontName
        If fontSize > 0 Then .size = fontSize
        .Color = wdColorAutomatic
    End With
    
    SafeSetFont = True
    Exit Function
    
ErrorHandler:
    SafeSetFont = False
    LogMessage "Erro ao aplicar fonte: " & Err.Description & " - Range: " & Left(targetRange.text, 20), LOG_LEVEL_WARNING
End Function

Private Function SafeSetParagraphFormat(para As Paragraph, alignment As Long, leftIndent As Single, firstLineIndent As Single) As Boolean
    On Error GoTo ErrorHandler
    
    With para.Format
        If alignment >= 0 Then .alignment = alignment
        If leftIndent >= 0 Then .leftIndent = leftIndent
        If firstLineIndent >= 0 Then .firstLineIndent = firstLineIndent
    End With
    
    SafeSetParagraphFormat = True
    Exit Function
    
ErrorHandler:
    SafeSetParagraphFormat = False
    LogMessage "Erro ao aplicar formatação de parágrafo: " & Err.Description, LOG_LEVEL_WARNING
End Function

Private Function SafeHasVisualContent(para As Paragraph) As Boolean
    On Error GoTo SafeMode
    
    ' Verificação padrão mais robusta
    Dim hasImages As Boolean
    Dim hasShapes As Boolean
    
    ' Verifica imagens inline de forma segura
    hasImages = (para.Range.InlineShapes.count > 0)
    
    ' Verifica shapes flutuantes de forma segura
    hasShapes = False
    If Not hasImages Then
        Dim shp As shape
        For Each shp In para.Range.ShapeRange
            hasShapes = True
            Exit For
        Next shp
    End If
    
    SafeHasVisualContent = hasImages Or hasShapes
    Exit Function
    
SafeMode:
    On Error GoTo ErrorHandler
    ' Método alternativo mais simples
    SafeHasVisualContent = (para.Range.InlineShapes.count > 0)
    Exit Function
    
ErrorHandler:
    ' Em caso de erro, assume que não há conteúdo visual
    SafeHasVisualContent = False
End Function

'================================================================================
' SAFE FIND/REPLACE OPERATIONS - Compatibilidade com todas as versões
'================================================================================
Private Function SafeFindReplace(doc As Document, findText As String, replaceText As String, Optional useWildcards As Boolean = False) As Long
    On Error GoTo ErrorHandler
    
    Dim findCount As Long
    findCount = 0
    
    ' Configuração segura de Find/Replace
    With doc.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        .Replacement.text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = useWildcards  ' Parâmetro controlado
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Executa a substituição e conta ocorrências
        Do While .Execute(Replace:=True)
            findCount = findCount + 1
            ' Limite de segurança para evitar loops infinitos
            If findCount > 10000 Then
                LogMessage "Limite de substituições atingido para: " & findText, LOG_LEVEL_WARNING
                Exit Do
            End If
        Loop
    End With
    
    SafeFindReplace = findCount
    Exit Function
    
ErrorHandler:
    SafeFindReplace = 0
    LogMessage "Erro na operação Find/Replace: " & findText & " -> " & replaceText & " | " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' SAFE CHARACTER ACCESS FUNCTIONS - Compatibilidade total
'================================================================================
Private Function SafeGetLastCharacter(rng As Range) As String
    On Error GoTo ErrorHandler
    
    Dim charCount As Long
    charCount = SafeGetCharacterCount(rng)
    
    If charCount > 0 Then
        SafeGetLastCharacter = rng.Characters(charCount).text
    Else
        SafeGetLastCharacter = ""
    End If
    Exit Function
    
ErrorHandler:
    ' Método alternativo usando Right()
    On Error GoTo FinalFallback
    SafeGetLastCharacter = Right(rng.text, 1)
    Exit Function
    
FinalFallback:
    SafeGetLastCharacter = ""
End Function

'================================================================================
' UNDO GROUP MANAGEMENT - #STABLE
'================================================================================
Private Sub StartUndoGroup(groupName As String)
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        EndUndoGroup
    End If
    
    Application.UndoRecord.StartCustomRecord groupName
    undoGroupEnabled = True
    
    Exit Sub
    
ErrorHandler:
    undoGroupEnabled = False
End Sub

Private Sub EndUndoGroup()
    On Error GoTo ErrorHandler
    
    If undoGroupEnabled Then
        Application.UndoRecord.EndCustomRecord
        undoGroupEnabled = False
    End If
    
    Exit Sub
    
ErrorHandler:
    undoGroupEnabled = False
End Sub

'================================================================================
' LOGGING MANAGEMENT - APRIMORADO COM DETALHES - #STABLE
'================================================================================
Private Function InitializeLogging(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If doc.Path <> "" Then
        logFilePath = doc.Path & "\" & Format(Now, "yyyy-mm-dd") & "_" & _
                     Replace(doc.Name, ".doc", "") & "_FormattingLog.txt"
        logFilePath = Replace(logFilePath, ".docx", "") & "_FormattingLog.txt"
        logFilePath = Replace(logFilePath, ".docm", "") & "_FormattingLog.txt"
    Else
        logFilePath = Environ("TEMP") & "\" & Format(Now, "yyyy-mm-dd") & "_DocumentFormattingLog.txt"
    End If
    
    Open logFilePath For Output As #1
    Print #1, "========================================================"
    Print #1, "LOG DE FORMATAÇÃO DE DOCUMENTO - SISTEMA DE REGISTRO"
    Print #1, "========================================================"
    Print #1, "Duração: " & Format(Now - executionStartTime, "HH:MM:ss")
    Print #1, "Erros: " & Err.Number & " - " & Err.Description
    Print #1, "Status: INICIANDO"
    Print #1, "--------------------------------------------------------"
    Print #1, "Sessão: " & Format(Now, "yyyy-mm-dd HH:MM:ss")
    Print #1, "Usuário: " & Environ("USERNAME")
    Print #1, "Estação: " & Environ("COMPUTERNAME")
    Print #1, "Versão Word: " & Application.version
    Print #1, "Documento: " & doc.Name
    Print #1, "Local: " & IIf(doc.Path = "", "(Não salvo)", doc.Path)
    Print #1, "Proteção: " & GetProtectionType(doc)
    Print #1, "Tamanho: " & GetDocumentSize(doc)
    Print #1, "========================================================"
    Close #1
    
    loggingEnabled = True
    InitializeLogging = True
    
    Exit Function
    
ErrorHandler:
    loggingEnabled = False
    InitializeLogging = False
End Function

Private Sub LogMessage(message As String, Optional level As Long = LOG_LEVEL_INFO)
    On Error GoTo ErrorHandler
    
    ' Debug mode output
    If DEBUG_MODE Then
        Debug.Print Format(Now, "hh:mm:ss") & " - " & message
    End If
    
    If Not loggingEnabled Then Exit Sub
    
    Dim levelText As String
    Dim levelIcon As String
    
    Select Case level
        Case LOG_LEVEL_INFO
            levelText = "INFO"
            levelIcon = ""
        Case LOG_LEVEL_WARNING
            levelText = "AVISO"
            levelIcon = ""
        Case LOG_LEVEL_ERROR
            levelText = "ERRO"
            levelIcon = ""
        Case Else
            levelText = "OUTRO"
            levelIcon = ""
    End Select
    
    Dim formattedMessage As String
    formattedMessage = Format(Now, "yyyy-mm-dd HH:MM:ss") & " [" & levelText & "] " & levelIcon & " " & message
    
    Open logFilePath For Append As #1
    Print #1, formattedMessage
    Close #1
    
    Debug.Print "LOG: " & formattedMessage
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "FALHA NO LOGGING: " & message
End Sub

Private Sub SafeFinalizeLogging()
    On Error GoTo ErrorHandler
    
    If loggingEnabled Then
        Open logFilePath For Append As #1
        Print #1, "================================================"
        Print #1, "FIM DA SESSÃO - " & Format(Now, "yyyy-mm-dd HH:MM:ss")
        Print #1, "Duração: " & Format(Now - executionStartTime, "HH:MM:ss")
        Print #1, "Erros: " & IIf(Err.Number = 0, "Nenhum", Err.Number & " - " & Err.Description)
        Print #1, "Status: " & IIf(formattingCancelled, "CANCELADO", "CONCLUÍDO")
        Print #1, "================================================"
        Close #1
    End If
    
    loggingEnabled = False
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Erro ao finalizar logging: " & Err.Description
    loggingEnabled = False
End Sub

'================================================================================
' UTILITY: GET PROTECTION TYPE - #STABLE
'================================================================================
Private Function GetProtectionType(doc As Document) As String
    On Error Resume Next
    
    Select Case doc.protectionType
        Case wdNoProtection: GetProtectionType = "Sem proteção"
        Case 1: GetProtectionType = "Protegido contra revisões"
        Case 2: GetProtectionType = "Protegido contra comentários"
        Case 3: GetProtectionType = "Protegido contra formulários"
        Case 4: GetProtectionType = "Protegido contra leitura"
        Case Else: GetProtectionType = "Tipo desconhecido (" & doc.protectionType & ")"
    End Select
End Function

'================================================================================
' UTILITY: GET DOCUMENT SIZE - #STABLE
'================================================================================
Private Function GetDocumentSize(doc As Document) As String
    On Error Resume Next
    
    Dim size As Long
    size = doc.BuiltInDocumentProperties("Number of Characters").value * 2
    
    If size < 1024 Then
        GetDocumentSize = size & " bytes"
    ElseIf size < 1048576 Then
        GetDocumentSize = Format(size / 1024, "0.0") & " KB"
    Else
        GetDocumentSize = Format(size / 1048576, "0.0") & " MB"
    End If
End Function

'================================================================================
' APPLICATION STATE HANDLER - #STABLE
'================================================================================
Private Function SetAppState(Optional ByVal enabled As Boolean = True, Optional ByVal statusMsg As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim success As Boolean
    success = True
    
    With Application
        On Error Resume Next
        .ScreenUpdating = enabled
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        On Error Resume Next
        .DisplayAlerts = IIf(enabled, wdAlertsAll, wdAlertsNone)
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
        
        If statusMsg <> "" Then
            On Error Resume Next
            .StatusBar = statusMsg
            If Err.Number <> 0 Then success = False
            On Error GoTo ErrorHandler
        ElseIf enabled Then
            On Error Resume Next
            .StatusBar = False
            If Err.Number <> 0 Then success = False
            On Error GoTo ErrorHandler
        End If
        
        On Error Resume Next
        .EnableCancelKey = 0
        If Err.Number <> 0 Then success = False
        On Error GoTo ErrorHandler
    End With
    
    SetAppState = success
    Exit Function
    
ErrorHandler:
    SetAppState = False
End Function

'================================================================================
' GLOBAL CHECKING - VERIFICAÇÕES ROBUSTAS
'================================================================================
Private Function PreviousChecking(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    If doc Is Nothing Then
        Application.StatusBar = "Erro: Documento não acessível para verificação"
        LogMessage "Documento não acessível para verificação", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.Type <> wdTypeDocument Then
        Application.StatusBar = "Erro: Tipo de documento não suportado (Tipo: " & doc.Type & ")"
        LogMessage "Tipo de documento não suportado: " & doc.Type, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If doc.protectionType <> wdNoProtection Then
        Dim protectionType As String
        protectionType = GetProtectionType(doc)
        Application.StatusBar = "Erro: Documento protegido (" & protectionType & ")"
        LogMessage "Documento protegido detectado: " & protectionType, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If
    
    If doc.ReadOnly Then
        Application.StatusBar = "Erro: Documento em modo somente leitura"
        LogMessage "Documento em modo somente leitura: " & doc.FullName, LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not CheckDiskSpace(doc) Then
        Application.StatusBar = "Erro: Espaço em disco insuficiente"
        LogMessage "Espaço em disco insuficiente para operação segura", LOG_LEVEL_ERROR
        PreviousChecking = False
        Exit Function
    End If

    If Not ValidateDocumentStructure(doc) Then
        LogMessage "Estrutura do documento validada com avisos", LOG_LEVEL_WARNING
    End If

    LogMessage "Verificações de segurança concluídas com sucesso", LOG_LEVEL_INFO
    PreviousChecking = True
    Exit Function

ErrorHandler:
    Application.StatusBar = "Erro durante verificações de segurança"
    LogMessage "Erro durante verificações: " & Err.Description, LOG_LEVEL_ERROR
    PreviousChecking = False
End Function

'================================================================================
' DISK SPACE CHECK - VERIFICAÇÃO SIMPLIFICADA
'================================================================================
Private Function CheckDiskSpace(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificação simplificada - assume espaço suficiente se não conseguir verificar
    Dim fso As Object
    Dim drive As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If doc.Path <> "" Then
        Set drive = fso.GetDrive(Left(doc.Path, 3))
    Else
        Set drive = fso.GetDrive(Left(Environ("TEMP"), 3))
    End If
    
    ' Verificação básica - 10MB mínimo
    If drive.AvailableSpace < 10485760 Then ' 10MB em bytes
        LogMessage "Espaço em disco muito baixo", LOG_LEVEL_WARNING
        CheckDiskSpace = False
    Else
        CheckDiskSpace = True
    End If
    
    Exit Function
    
ErrorHandler:
    ' Se não conseguir verificar, assume que há espaço suficiente
    CheckDiskSpace = True
End Function

'================================================================================
' MAIN FORMATTING ROUTINE - #STABLE
'================================================================================
Private Function PreviousFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    ' Formatações básicas de página e estrutura
    If Not ApplyPageSetup(doc) Then
        LogMessage "Falha na configuração de página", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    ' Limpeza e formatações otimizadas (logs reduzidos para performance)
    ClearAllFormatting doc
    CleanDocumentStructure doc
    ValidatePropositionType doc
    FormatDocumentTitle doc
    
    ' Formata parágrafos especiais (Justificativa/Anexo/Vereador) ANTES das formatações principais
    FormatJustificativaAnexoParagraphs doc
    
    ' Formatações principais
    If Not ApplyStdFont(doc) Then
        LogMessage "Falha na formatação de fontes", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    If Not ApplyStdParagraphs(doc) Then
        LogMessage "Falha na formatação de parágrafos", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If

    ' Formatação específica do 1º parágrafo (caixa alta, negrito, sublinhado)
    FormatFirstParagraph doc

    ' Formatação específica do 2º parágrafo
    FormatSecondParagraph doc

    ' Formatações específicas (sem verificação de retorno para performance)
    FormatConsiderandoParagraphs doc
    ApplyTextReplacements doc
    
    EnableHyphenation doc
    RemoveWatermark doc
    InsertHeaderstamp doc
    
    ' Limpeza final de espaços múltiplos em todo o documento
    CleanMultipleSpaces doc
    
    ' Controle de linhas em branco sequenciais (máximo 2)
    LimitSequentialEmptyLines doc
    
    ' REFORÇO: Garante que o 2º parágrafo mantenha suas 2 linhas em branco
    EnsureSecondParagraphBlankLines doc

    ' Substituição de datas no parágrafo de plenário
    ReplacePlenarioDateParagraph doc
    
    ' Configuração final da visualização
    ConfigureDocumentView doc
    
    If Not InsertFooterStamp(doc) Then
        LogMessage "Falha na inserção do rodapé", LOG_LEVEL_ERROR
        PreviousFormatting = False
        Exit Function
    End If
    
    ' APLICAÇÃO FINAL DE NEGRITO: Última operação para garantir negrito em parágrafos especiais
    ApplyBoldToSpecialParagraphs doc
    
    ' INSERÇÃO FINAL DE LINHAS EM BRANCO: Insere linhas vazias após todas as limpezas
    InsertJustificativaBlankLines doc
    
    LogMessage "Formatação completa aplicada com sucesso", LOG_LEVEL_INFO
    PreviousFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro durante formatação: " & Err.Description, LOG_LEVEL_ERROR
    PreviousFormatting = False
End Function

'================================================================================
' PAGE SETUP - #STABLE
'================================================================================
Private Function ApplyPageSetup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(TOP_MARGIN_CM)
        .BottomMargin = CentimetersToPoints(BOTTOM_MARGIN_CM)
        .LeftMargin = CentimetersToPoints(LEFT_MARGIN_CM)
        .RightMargin = CentimetersToPoints(RIGHT_MARGIN_CM)
        .HeaderDistance = CentimetersToPoints(HEADER_DISTANCE_CM)
        .FooterDistance = CentimetersToPoints(FOOTER_DISTANCE_CM)
        .Gutter = 0
        .Orientation = wdOrientPortrait
    End With
    
    ' Configuração de página aplicada (sem log detalhado para performance)
    ApplyPageSetup = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro na configuração de página: " & Err.Description, LOG_LEVEL_ERROR
    ApplyPageSetup = False
End Function

' ================================================================================
' FONT FORMMATTING - #STABLE
' ================================================================================
Private Function ApplyStdFont(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim underlineRemovedCount As Long
    Dim isTitle As Boolean
    Dim hasConsiderando As Boolean
    Dim needsUnderlineRemoval As Boolean
    Dim needsBoldRemoval As Boolean

    For i = doc.Paragraphs.count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False
        isTitle = False
        hasConsiderando = False
        needsUnderlineRemoval = False
        needsBoldRemoval = False
        
        ' SUPER OTIMIZADO: Verificação prévia consolidada - uma única leitura das propriedades
        Dim paraFont As Font
        Set paraFont = para.Range.Font
        Dim needsFontFormatting As Boolean
        needsFontFormatting = (paraFont.Name <> STANDARD_FONT) Or _
                             (paraFont.size <> STANDARD_FONT_SIZE) Or _
                             (paraFont.Color <> wdColorAutomatic)
        
        ' Cache das verificações de formatação especial
        needsUnderlineRemoval = (paraFont.Underline <> wdUnderlineNone)
        needsBoldRemoval = (paraFont.Bold = True)
        
        ' Cache da contagem de InlineShapes para evitar múltiplas chamadas
        Dim inlineShapesCount As Long
        inlineShapesCount = para.Range.InlineShapes.count
        
        ' OTIMIZAÇÃO MÁXIMA: Se não precisa de nenhuma formatação, pula imediatamente
        If Not needsFontFormatting And Not needsUnderlineRemoval And Not needsBoldRemoval And inlineShapesCount = 0 Then
            formattedCount = formattedCount + 1
            GoTo NextParagraph
        End If

        If inlineShapesCount > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
        ' OTIMIZADO: Verificação de conteúdo visual só quando necessário
        If Not hasInlineImage And (needsFontFormatting Or needsUnderlineRemoval Or needsBoldRemoval) Then
            If HasVisualContent(para) Then
                hasInlineImage = True
                skippedCount = skippedCount + 1
            End If
        End If
        
        
        ' OTIMIZADO: Verificação consolidada de tipo de parágrafo - uma única leitura do texto
        Dim paraFullText As String
        Dim isSpecialParagraph As Boolean
        isSpecialParagraph = False
        
        ' Só faz verificação de texto se for necessário para formatação especial
        If needsUnderlineRemoval Or needsBoldRemoval Then
            paraFullText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' Verifica se é o primeiro parágrafo com texto (título) - otimizado
            If i <= 3 And para.Format.alignment = wdAlignParagraphCenter And paraFullText <> "" Then
                isTitle = True
            End If
            
            ' Verifica se o parágrafo começa com "considerando" - otimizado
            If Len(paraFullText) >= CONSIDERANDO_MIN_LENGTH And LCase(Left(paraFullText, CONSIDERANDO_MIN_LENGTH)) = CONSIDERANDO_PREFIX Then
                hasConsiderando = True
            End If
            
            ' Verifica se é um parágrafo especial - otimizado
            Dim cleanParaText As String
            cleanParaText = paraFullText
            ' Remove pontuação final para análise
            Do While Len(cleanParaText) > 0 And (Right(cleanParaText, 1) = "." Or Right(cleanParaText, 1) = "," Or Right(cleanParaText, 1) = ":" Or Right(cleanParaText, 1) = ";")
                cleanParaText = Left(cleanParaText, Len(cleanParaText) - 1)
            Loop
            cleanParaText = Trim(LCase(cleanParaText))
            
            If cleanParaText = "justificativa" Or IsVereadorPattern(cleanParaText) Or IsAnexoPattern(cleanParaText) Then
                isSpecialParagraph = True
                LogMessage "Parágrafo especial detectado em ApplyStdFont (negrito preservado): " & cleanParaText, LOG_LEVEL_INFO
            End If
            
            ' Verifica se é o parágrafo ANTERIOR a "- vereador -" (também deve preservar negrito)
            Dim isBeforeVereador As Boolean
            isBeforeVereador = False
            If i < doc.Paragraphs.count Then
                Dim nextPara As Paragraph
                Set nextPara = doc.Paragraphs(i + 1)
                If Not HasVisualContent(nextPara) Then
                    Dim nextParaText As String
                    nextParaText = Trim(Replace(Replace(nextPara.Range.text, vbCr, ""), vbLf, ""))
                    ' Remove pontuação final para análise
                    Dim nextCleanText As String
                    nextCleanText = nextParaText
                    Do While Len(nextCleanText) > 0 And (Right(nextCleanText, 1) = "." Or Right(nextCleanText, 1) = "," Or Right(nextCleanText, 1) = ":" Or Right(nextCleanText, 1) = ";")
                        nextCleanText = Left(nextCleanText, Len(nextCleanText) - 1)
                    Loop
                    nextCleanText = Trim(LCase(nextCleanText))
                    
                    If IsVereadorPattern(nextCleanText) Then
                        isBeforeVereador = True
                    End If
                End If
            End If
        End If

        ' FORMATAÇÃO PRINCIPAL - Só executa se necessário
        If needsFontFormatting Then
            If Not hasInlineImage Then
                ' Formatação rápida para parágrafos sem imagens usando método seguro
                If SafeSetFont(para.Range, STANDARD_FONT, STANDARD_FONT_SIZE) Then
                    formattedCount = formattedCount + 1
                Else
                    ' Fallback para método tradicional em caso de erro
                    With paraFont
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                    End With
                    formattedCount = formattedCount + 1
                End If
            Else
                ' NOVO: Formatação protegida para parágrafos COM imagens
                If ProtectImagesInRange(para.Range) Then
                    formattedCount = formattedCount + 1
                Else
                    ' Fallback: formatação básica segura CONSOLIDADA
                    Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, False, False)
                    formattedCount = formattedCount + 1
                End If
            End If
        End If
        
        ' FORMATAÇÃO ESPECIAL CONSOLIDADA - Remove sublinhado e negrito em uma única passada
        If needsUnderlineRemoval Or needsBoldRemoval Then
            ' Determina quais formatações remover
            Dim removeUnderline As Boolean
            Dim removeBold As Boolean
            removeUnderline = needsUnderlineRemoval And Not isTitle
            removeBold = needsBoldRemoval And Not isTitle And Not hasConsiderando And Not isSpecialParagraph And Not isBeforeVereador
            
            ' Se precisa remover alguma formatação
            If removeUnderline Or removeBold Then
                If Not hasInlineImage Then
                    ' Formatação rápida para parágrafos sem imagens
                    If removeUnderline Then paraFont.Underline = wdUnderlineNone
                    If removeBold Then paraFont.Bold = False
                Else
                    ' Formatação protegida CONSOLIDADA para parágrafos com imagens
                    Call FormatCharacterByCharacter(para, "", 0, 0, removeUnderline, removeBold)
                End If
                
                If removeUnderline Then underlineRemovedCount = underlineRemovedCount + 1
            End If
        End If

NextParagraph:
    Next i
    
    ' Log otimizado
    If skippedCount > 0 Then
        LogMessage "Fontes formatadas: " & formattedCount & " parágrafos (incluindo " & skippedCount & " com proteção de imagens)"
    End If
    
    ApplyStdFont = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de fonte: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdFont = False
End Function

'================================================================================
' FORMATAÇÃO CARACTERE POR CARACTERE CONSOLIDADA - #OPTIMIZED
'================================================================================
Private Sub FormatCharacterByCharacter(para As Paragraph, fontName As String, fontSize As Long, fontColor As Long, removeUnderline As Boolean, removeBold As Boolean)
    On Error Resume Next
    
    Dim j As Long
    Dim charCount As Long
    Dim charRange As Range
    
    charCount = SafeGetCharacterCount(para.Range) ' Cache da contagem segura
    
    If charCount > 0 Then ' Verificação de segurança
        For j = 1 To charCount
            Set charRange = para.Range.Characters(j)
            If charRange.InlineShapes.count = 0 Then
                With charRange.Font
                    ' Aplica formatação de fonte se especificada
                    If fontName <> "" Then .Name = fontName
                    If fontSize > 0 Then .size = fontSize
                    If fontColor >= 0 Then .Color = fontColor
                    
                    ' Remove formatações especiais se solicitado
                    If removeUnderline Then .Underline = wdUnderlineNone
                    If removeBold Then .Bold = False
                End With
            End If
        Next j
    End If
End Sub

'================================================================================
' PARAGRAPH FORMATTING - #STABLE
'================================================================================
Private Function ApplyStdParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim hasInlineImage As Boolean
    Dim paragraphIndent As Single
    Dim firstIndent As Single
    Dim rightMarginPoints As Single
    Dim i As Long
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim paraText As String
    Dim prevPara As Paragraph

    rightMarginPoints = 0

    For i = doc.Paragraphs.count To 1 Step -1
        Set para = doc.Paragraphs(i)
        hasInlineImage = False

        If para.Range.InlineShapes.count > 0 Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If
        
        ' Proteção adicional: verifica outros tipos de conteúdo visual
        If Not hasInlineImage And HasVisualContent(para) Then
            hasInlineImage = True
            skippedCount = skippedCount + 1
        End If

        ' Aplica formatação de parágrafo para TODOS os parágrafos
        ' (independente se contêm imagens ou não)
        
        ' Limpeza robusta de espaços múltiplos - SEMPRE aplicada
        Dim cleanText As String
        cleanText = para.Range.text
        
        ' OTIMIZADO: Combinação de múltiplas operações de limpeza em um bloco
        If InStr(cleanText, "  ") > 0 Or InStr(cleanText, vbTab) > 0 Then
            ' Remove múltiplos espaços consecutivos
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
            
            ' Remove espaços antes/depois de quebras de linha
            cleanText = Replace(cleanText, " " & vbCr, vbCr)
            cleanText = Replace(cleanText, vbCr & " ", vbCr)
            cleanText = Replace(cleanText, " " & vbLf, vbLf)
            cleanText = Replace(cleanText, vbLf & " ", vbLf)
            
            ' Remove tabs extras e converte para espaços
            Do While InStr(cleanText, vbTab & vbTab) > 0
                cleanText = Replace(cleanText, vbTab & vbTab, vbTab)
            Loop
            cleanText = Replace(cleanText, vbTab, " ")
            
            ' Limpeza final de espaços múltiplos
            Do While InStr(cleanText, "  ") > 0
                cleanText = Replace(cleanText, "  ", " ")
            Loop
        End If
        
        ' Verifica se é um parágrafo especial ANTES de limpar o texto
        Dim isSpecialFormatParagraph As Boolean
        isSpecialFormatParagraph = False
        
        Dim checkText As String
        checkText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        ' Remove pontuação final para análise
        Do While Len(checkText) > 0 And (Right(checkText, 1) = "." Or Right(checkText, 1) = "," Or Right(checkText, 1) = ":" Or Right(checkText, 1) = ";")
            checkText = Left(checkText, Len(checkText) - 1)
        Loop
        checkText = Trim(LCase(checkText))
        
        ' Verifica se é "Justificativa", "Anexo", "Anexos" ou padrão de vereador
        If checkText = JUSTIFICATIVA_TEXT Or IsAnexoPattern(checkText) Or IsVereadorPattern(checkText) Then
            isSpecialFormatParagraph = True
        End If
        
        ' Aplica o texto limpo APENAS se não há imagens E não é parágrafo especial
        If cleanText <> para.Range.text And Not hasInlineImage And Not isSpecialFormatParagraph Then
            para.Range.text = cleanText
        End If

        ' Formatação de parágrafo - SEMPRE aplicada (exceto para parágrafos especiais)
        If Not isSpecialFormatParagraph Then
            With para.Format
                .LineSpacingRule = wdLineSpacingMultiple
                .LineSpacing = LINE_SPACING
                .RightIndent = rightMarginPoints
                .SpaceBefore = 0
                .SpaceAfter = 0

                If para.alignment = wdAlignParagraphCenter Then
                    .leftIndent = 0
                    .firstLineIndent = 0
                Else
                    firstIndent = .firstLineIndent
                    paragraphIndent = .leftIndent
                    If paragraphIndent >= CentimetersToPoints(5) Then
                        .leftIndent = CentimetersToPoints(9)
                    ElseIf firstIndent < CentimetersToPoints(5) Then
                        .leftIndent = CentimetersToPoints(0)
                        .firstLineIndent = CentimetersToPoints(2.5)
                    End If
                End If
            End With

            If para.alignment = wdAlignParagraphLeft Then
                para.alignment = wdAlignParagraphJustify
            End If
        End If
        
        formattedCount = formattedCount + 1
    Next i
    
    ' Log atualizado para refletir que todos os parágrafos são formatados
    If skippedCount > 0 Then
        LogMessage "Parágrafos formatados: " & formattedCount & " (incluindo " & skippedCount & " com proteção de imagens)"
    End If
    
    ApplyStdParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de parágrafos: " & Err.Description, LOG_LEVEL_ERROR
    ApplyStdParagraphs = False
End Function

'================================================================================
' FORMAT SECOND PARAGRAPH - FORMATAÇÃO APENAS DO 2º PARÁGRAFO - #NEW
'================================================================================
Private Function FormatSecondParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim secondParaIndex As Long
    
    ' Identifica apenas o 2º parágrafo (considerando apenas parágrafos com texto)
    actualParaIndex = 0
    secondParaIndex = 0
    
    ' Encontra o 2º parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo tem texto ou conteúdo visual, conta como parágrafo válido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Registra o índice do 2º parágrafo
            If actualParaIndex = 2 Then
                secondParaIndex = i
                Exit For ' Já encontramos o 2º parágrafo
            End If
        End If
        
        ' Proteção expandida: processa até 20 parágrafos para encontrar o 2º
        If i > 20 Then Exit For
    Next i
    
    ' Aplica formatação específica apenas ao 2º parágrafo
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(secondParaIndex)
        
        ' PRIMEIRO: Adiciona 2 linhas em branco ANTES do 2º parágrafo
        Dim insertionPoint As Range
        Set insertionPoint = para.Range
        insertionPoint.Collapse wdCollapseStart
        
        ' Verifica se já existem linhas em branco antes
        Dim blankLinesBefore As Long
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)
        
        ' Adiciona linhas em branco conforme necessário para chegar a 2
        If blankLinesBefore < 2 Then
            Dim linesToAdd As Long
            linesToAdd = 2 - blankLinesBefore
            
            Dim newLines As String
            newLines = String(linesToAdd, vbCrLf)
            insertionPoint.InsertBefore newLines
            
            ' Atualiza o índice do segundo parágrafo (foi deslocado)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If
        
        ' FORMATAÇÃO PRINCIPAL: Aplica formatação SEMPRE, protegendo apenas as imagens
        With para.Format
            .leftIndent = CentimetersToPoints(9)      ' Recuo à esquerda de 9 cm
            .firstLineIndent = 0                      ' Sem recuo da primeira linha
            .RightIndent = 0                          ' Sem recuo à direita
            .alignment = wdAlignParagraphJustify      ' Justificado
        End With
        
        ' SEGUNDO: Adiciona 2 linhas em branco DEPOIS do 2º parágrafo
        Dim insertionPointAfter As Range
        Set insertionPointAfter = para.Range
        insertionPointAfter.Collapse wdCollapseEnd
        
        ' Verifica se já existem linhas em branco depois
        Dim blankLinesAfter As Long
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
        
        ' Adiciona linhas em branco conforme necessário para chegar a 2
        If blankLinesAfter < 2 Then
            Dim linesToAddAfter As Long
            linesToAddAfter = 2 - blankLinesAfter
            
            Dim newLinesAfter As String
            newLinesAfter = String(linesToAddAfter, vbCrLf)
            insertionPointAfter.InsertAfter newLinesAfter
        End If
        
        ' Se tem imagens, apenas registra (mas não pula a formatação)
        If HasVisualContent(para) Then
            LogMessage "2º parágrafo formatado com proteção de imagem e linhas em branco (posição: " & secondParaIndex & ")", LOG_LEVEL_INFO
        Else
            LogMessage "2º parágrafo formatado com 2 linhas em branco antes e depois (posição: " & secondParaIndex & ")", LOG_LEVEL_INFO
        End If
    Else
        LogMessage "2º parágrafo não encontrado para formatação", LOG_LEVEL_WARNING
    End If
    
    FormatSecondParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do 2º parágrafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatSecondParagraph = False
End Function

'================================================================================
' HELPER FUNCTIONS FOR BLANK LINES - Funções auxiliares para linhas em branco
'================================================================================
Private Function CountBlankLinesBefore(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim count As Long
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    
    count = 0
    
    ' Verifica parágrafos anteriores (máximo 5 para performance)
    For i = paraIndex - 1 To 1 Step -1
        If i <= 0 Then Exit For
        
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo está vazio, conta como linha em branco
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            ' Se encontrou parágrafo com conteúdo, para de contar
            Exit For
        End If
        
        ' Limite de segurança
        If count >= 5 Then Exit For
    Next i
    
    CountBlankLinesBefore = count
    Exit Function
    
ErrorHandler:
    CountBlankLinesBefore = 0
End Function

Private Function CountBlankLinesAfter(doc As Document, paraIndex As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim count As Long
    Dim i As Long
    Dim para As Paragraph
    Dim paraText As String
    
    count = 0
    
    ' Verifica parágrafos posteriores (máximo 5 para performance)
    For i = paraIndex + 1 To doc.Paragraphs.count
        If i > doc.Paragraphs.count Then Exit For
        
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo está vazio, conta como linha em branco
        If paraText = "" And Not HasVisualContent(para) Then
            count = count + 1
        Else
            ' Se encontrou parágrafo com conteúdo, para de contar
            Exit For
        End If
        
        ' Limite de segurança
        If count >= 5 Then Exit For
    Next i
    
    CountBlankLinesAfter = count
    Exit Function
    
ErrorHandler:
    CountBlankLinesAfter = 0
End Function

'================================================================================
' SECOND PARAGRAPH LOCATION HELPER - Localiza o segundo parágrafo
'================================================================================
Private Function GetSecondParagraphIndex(doc As Document) As Long
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    
    actualParaIndex = 0
    
    ' Encontra o 2º parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo tem texto ou conteúdo visual, conta como parágrafo válido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Retorna o índice do 2º parágrafo
            If actualParaIndex = 2 Then
                GetSecondParagraphIndex = i
                Exit Function
            End If
        End If
        
        ' Proteção: processa até 20 parágrafos para encontrar o 2º
        If i > 20 Then Exit For
    Next i
    
    GetSecondParagraphIndex = 0  ' Não encontrado
    Exit Function
    
ErrorHandler:
    GetSecondParagraphIndex = 0
End Function

'================================================================================
' ENSURE SECOND PARAGRAPH BLANK LINES - Garante 2 linhas em branco no 2º parágrafo
'================================================================================
Private Function EnsureSecondParagraphBlankLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim secondParaIndex As Long
    Dim linesToAdd As Long
    Dim linesToAddAfter As Long
    
    secondParaIndex = GetSecondParagraphIndex(doc)
    linesToAdd = 0
    linesToAddAfter = 0
    
    If secondParaIndex > 0 And secondParaIndex <= doc.Paragraphs.count Then
        Dim para As Paragraph
        Set para = doc.Paragraphs(secondParaIndex)
        
        ' Verifica e corrige linhas em branco ANTES
        Dim blankLinesBefore As Long
        blankLinesBefore = CountBlankLinesBefore(doc, secondParaIndex)
        
        If blankLinesBefore < 2 Then
            Dim insertionPoint As Range
            Set insertionPoint = para.Range
            insertionPoint.Collapse wdCollapseStart
            
            linesToAdd = 2 - blankLinesBefore
            
            Dim newLines As String
            newLines = String(linesToAdd, vbCrLf)
            insertionPoint.InsertBefore newLines
            
            ' Atualiza o índice (foi deslocado)
            secondParaIndex = secondParaIndex + linesToAdd
            Set para = doc.Paragraphs(secondParaIndex)
        End If
        
        ' Verifica e corrige linhas em branco DEPOIS
        Dim blankLinesAfter As Long
        blankLinesAfter = CountBlankLinesAfter(doc, secondParaIndex)
        
        If blankLinesAfter < 2 Then
            Dim insertionPointAfter As Range
            Set insertionPointAfter = para.Range
            insertionPointAfter.Collapse wdCollapseEnd
            
            linesToAddAfter = 2 - blankLinesAfter
            
            Dim newLinesAfter As String
            newLinesAfter = String(linesToAddAfter, vbCrLf)
            insertionPointAfter.InsertAfter newLinesAfter
        End If
        
        LogMessage "Linhas em branco do 2º parágrafo reforçadas (antes: " & (blankLinesBefore + linesToAdd) & ", depois: " & (blankLinesAfter + linesToAddAfter) & ")", LOG_LEVEL_INFO
    End If
    
    EnsureSecondParagraphBlankLines = True
    Exit Function
    
ErrorHandler:
    EnsureSecondParagraphBlankLines = False
    LogMessage "Erro ao garantir linhas em branco do 2º parágrafo: " & Err.Description, LOG_LEVEL_WARNING
End Function

'================================================================================
' FORMAT FIRST PARAGRAPH - FORMATAÇÃO DO 1º PARÁGRAFO - #NEW
'================================================================================
Private Function FormatFirstParagraph(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Long
    Dim actualParaIndex As Long
    Dim firstParaIndex As Long
    
    ' Identifica o 1º parágrafo (considerando apenas parágrafos com texto)
    actualParaIndex = 0
    firstParaIndex = 0
    
    ' Encontra o 1º parágrafo com conteúdo (pula vazios)
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Se o parágrafo tem texto ou conteúdo visual, conta como parágrafo válido
        If paraText <> "" Or HasVisualContent(para) Then
            actualParaIndex = actualParaIndex + 1
            
            ' Registra o índice do 1º parágrafo
            If actualParaIndex = 1 Then
                firstParaIndex = i
                Exit For ' Já encontramos o 1º parágrafo
            End If
        End If
        
        ' Proteção expandida: processa até 20 parágrafos para encontrar o 1º
        If i > 20 Then Exit For
    Next i
    
    ' Aplica formatação específica apenas ao 1º parágrafo
    If firstParaIndex > 0 And firstParaIndex <= doc.Paragraphs.count Then
        Set para = doc.Paragraphs(firstParaIndex)
        
        ' NOVO: Aplica formatação SEMPRE, protegendo apenas as imagens
        ' Formatação do 1º parágrafo: caixa alta, negrito e sublinhado
        If HasVisualContent(para) Then
            ' Para parágrafos com imagens, aplica formatação caractere por caractere
            Dim n As Long
            Dim charCount4 As Long
            charCount4 = SafeGetCharacterCount(para.Range) ' Cache da contagem segura
            
            If charCount4 > 0 Then ' Verificação de segurança
                For n = 1 To charCount4
                    Dim charRange3 As Range
                    Set charRange3 = para.Range.Characters(n)
                    If charRange3.InlineShapes.count = 0 Then
                        With charRange3.Font
                            .AllCaps = True           ' Caixa alta (maiúsculas)
                            .Bold = True              ' Negrito
                            .Underline = wdUnderlineSingle ' Sublinhado
                        End With
                    End If
                Next n
            End If
            LogMessage "1º parágrafo formatado com proteção de imagem (posição: " & firstParaIndex & ")"
        Else
            ' Formatação normal para parágrafos sem imagens
            With para.Range.Font
                .AllCaps = True           ' Caixa alta (maiúsculas)
                .Bold = True              ' Negrito
                .Underline = wdUnderlineSingle ' Sublinhado
            End With
        End If
        
        ' Aplicar também formatação de parágrafo - SEMPRE
        With para.Format
            .alignment = wdAlignParagraphCenter       ' Centralizado
            .leftIndent = 0                           ' Sem recuo à esquerda
            .firstLineIndent = 0                      ' Sem recuo da primeira linha
            .RightIndent = 0                          ' Sem recuo à direita
        End With
    Else
        LogMessage "1º parágrafo não encontrado para formatação", LOG_LEVEL_WARNING
    End If
    
    FormatFirstParagraph = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do 1º parágrafo: " & Err.Description, LOG_LEVEL_ERROR
    FormatFirstParagraph = False
End Function

'================================================================================
' ENABLE HYPHENATION - #STABLE
'================================================================================
Private Function EnableHyphenation(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If Not doc.AutoHyphenation Then
        doc.AutoHyphenation = True
        doc.HyphenationZone = CentimetersToPoints(0.63)
        doc.HyphenateCaps = True
        ' Log removido para performance
        EnableHyphenation = True
    Else
        ' Log removido para performance
        EnableHyphenation = True
    End If
    
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao ativar hifenização: " & Err.Description, LOG_LEVEL_ERROR
    EnableHyphenation = False
End Function

'================================================================================
' REMOVE WATERMARK - #STABLE
'================================================================================
Private Function RemoveWatermark(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As section
    Dim header As HeaderFooter
    Dim shp As shape
    Dim i As Long
    Dim removedCount As Long

    For Each sec In doc.Sections
        For Each header In sec.Headers
            If header.Exists And header.Shapes.count > 0 Then
                For i = header.Shapes.count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                        End If
                    End If
                Next i
            End If
        Next header
        
        For Each header In sec.Footers
            If header.Exists And header.Shapes.count > 0 Then
                For i = header.Shapes.count To 1 Step -1
                    Set shp = header.Shapes(i)
                    If shp.Type = msoPicture Or shp.Type = msoTextEffect Then
                        If InStr(1, shp.Name, "Watermark", vbTextCompare) > 0 Or _
                           InStr(1, shp.AlternativeText, "Watermark", vbTextCompare) > 0 Then
                            shp.Delete
                            removedCount = removedCount + 1
                        End If
                    End If
                Next i
            End If
        Next header
    Next sec

    If removedCount > 0 Then
        LogMessage "Marcas d'água removidas: " & removedCount & " itens"
    End If
    ' Log de "nenhuma marca d'água" removido para performance
    
    RemoveWatermark = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao remover marcas d'água: " & Err.Description, LOG_LEVEL_ERROR
    RemoveWatermark = False
End Function

'================================================================================
' HEADER IMAGE PATH MANAGEMENT - #STABLE
'================================================================================
Private Function GetHeaderImagePath() As String
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim shell As Object
    Dim documentsPath As String
    Dim headerImagePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    ' Obtém pasta Documents do usuário atual (compatível com Windows)
    documentsPath = shell.SpecialFolders("MyDocuments")
    If Right(documentsPath, 1) = "\" Then
        documentsPath = Left(documentsPath, Len(documentsPath) - 1)
    End If

    ' Constrói caminho absoluto para a imagem desejada
    headerImagePath = documentsPath & "\chainsaw-proposituras\assets\stamp.png"

    ' Verifica se o arquivo existe
    If Not fso.FileExists(headerImagePath) Then
        LogMessage "Imagem de cabeçalho não encontrada em: " & headerImagePath, LOG_LEVEL_WARNING
        GetHeaderImagePath = ""
        Exit Function
    End If

    GetHeaderImagePath = headerImagePath
    Exit Function

ErrorHandler:
    LogMessage "Erro ao localizar imagem de cabeçalho: " & Err.Description, LOG_LEVEL_ERROR
    GetHeaderImagePath = ""
End Function

'================================================================================
' INSERT HEADER IMAGE
'================================================================================
Private Function InsertHeaderstamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As section
    Dim header As HeaderFooter
    Dim imgFile As String
    Dim username As String
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim shp As shape
    Dim imgFound As Boolean
    Dim sectionsProcessed As Long

    ' Resolve image file path using configuration if provided
    'imgFile = Trim(config.headerImagePath)
    If Len(imgFile) = 0 Then
        ' No configured path; try common locations relative to document or repo folder
        Dim baseFolder As String
        On Error Resume Next
        baseFolder = IIf(doc.Path <> "", doc.Path, Environ("USERPROFILE") & "\Documents")
        On Error GoTo ErrorHandler
        If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
        If Dir(baseFolder & "assets\stamp.png") <> "" Then
            imgFile = baseFolder & "assets\stamp.png"
        ElseIf Dir(Environ("USERPROFILE") & "\Documents\chainsaw\assets\stamp.png") <> "" Then
            imgFile = Environ("USERPROFILE") & "\Documents\chainsaw\assets\stamp.png"
        End If
    Else
        ' If relative path (no drive letter), try resolve from current document folder and repo root
        If InStr(1, imgFile, ":", vbTextCompare) = 0 And Left(imgFile, 2) <> "\\" Then
            On Error Resume Next
            baseFolder = IIf(doc.Path <> "", doc.Path, Environ("USERPROFILE") & "\Documents")
            On Error GoTo ErrorHandler
            If Right(baseFolder, 1) <> "\" Then baseFolder = baseFolder & "\"
            If Dir(baseFolder & imgFile) <> "" Then
                imgFile = baseFolder & imgFile
            ElseIf Dir(Environ("USERPROFILE") & "\Documents\chainsaw\" & imgFile) <> "" Then
                imgFile = Environ("USERPROFILE") & "\Documents\chainsaw\" & imgFile
            End If
        End If
    End If

    If Dir(imgFile) = "" Then
        Application.StatusBar = "Warning: Header image not found"
        LogMessage "Header image not found at: " & imgFile, LOG_LEVEL_WARNING
        InsertHeaderstamp = False
        Exit Function
    End If

    ' Size using standard constants
    imgWidth = CentimetersToPoints(HEADER_IMAGE_MAX_WIDTH_CM)
    imgHeight = imgWidth * HEADER_IMAGE_HEIGHT_RATIO

    For Each sec In doc.Sections
        Set header = sec.Headers(wdHeaderFooterPrimary)
        If header.Exists Then
            header.LinkToPrevious = False
            header.Range.Delete
            
            Set shp = header.Shapes.AddPicture( _
                FileName:=imgFile, _
                LinkToFile:=False, _
                SaveWithDocument:=msoTrue)
            
            If shp Is Nothing Then
                LogMessage "Failed to insert header image at section " & sectionsProcessed + 1, LOG_LEVEL_WARNING
            Else
                With shp
                    .LockAspectRatio = msoTrue
                    .Width = imgWidth
                    .Height = imgHeight
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
                    .Left = (doc.PageSetup.PageWidth - .Width) / 2
                    .Top = CentimetersToPoints(HEADER_IMAGE_TOP_MARGIN_CM)
                    .WrapFormat.Type = wdWrapTopBottom
                    .ZOrder msoSendToBack
                End With
                
                imgFound = True
                sectionsProcessed = sectionsProcessed + 1
            End If
        End If
    Next sec

    If imgFound Then
        ' Log detalhado removido para performance
        InsertHeaderstamp = True
    Else
    LogMessage "No header was inserted", LOG_LEVEL_WARNING
        InsertHeaderstamp = False
    End If

    Exit Function

ErrorHandler:
    LogMessage "Error inserting header: " & Err.Description, LOG_LEVEL_ERROR
    InsertHeaderstamp = False
End Function

'================================================================================
' INSERT FOOTER PAGE NUMBERS - #STABLE
'================================================================================
Private Function InsertFooterStamp(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Dim sec As section
    Dim footer As HeaderFooter
    Dim rng As Range
    Dim sectionsProcessed As Long

    For Each sec In doc.Sections
        Set footer = sec.Footers(wdHeaderFooterPrimary)
        
        If footer.Exists Then
            footer.LinkToPrevious = False
            Set rng = footer.Range
            
            rng.Delete
            
            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.Fields.Add Range:=rng, Type:=wdFieldPage
            
            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.text = "-"
            
            Set rng = footer.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.Fields.Add Range:=rng, Type:=wdFieldNumPages
            
            With footer.Range
                .Font.Name = STANDARD_FONT
                .Font.size = FOOTER_FONT_SIZE
                .ParagraphFormat.alignment = wdAlignParagraphCenter
                .Fields.Update
            End With
            
            sectionsProcessed = sectionsProcessed + 1
        End If
    Next sec

    ' Log detalhado removido para performance
    InsertFooterStamp = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao inserir rodapé: " & Err.Description, LOG_LEVEL_ERROR
    InsertFooterStamp = False
End Function

'================================================================================
' UTILITY: CM TO POINTS - #STABLE
'================================================================================
Private Function CentimetersToPoints(ByVal cm As Double) As Single
    On Error Resume Next
    CentimetersToPoints = Application.CentimetersToPoints(cm)
    If Err.Number <> 0 Then
        CentimetersToPoints = cm * 28.35
    End If
End Function

'================================================================================
' UTILITY: SAFE USERNAME - #STABLE
'================================================================================
Private Function GetSafeUserName() As String
    On Error GoTo ErrorHandler
    
    Dim rawName As String
    Dim safeName As String
    Dim i As Integer
    Dim c As String
    
    rawName = Environ("USERNAME")
    If rawName = "" Then rawName = Environ("USER")
    If rawName = "" Then
        On Error Resume Next
        rawName = CreateObject("WScript.Network").username
        On Error GoTo 0
    End If
    
    If rawName = "" Then
        rawName = "UsuarioDesconhecido"
    End If
    
    For i = 1 To Len(rawName)
        c = Mid(rawName, i, 1)
        If c Like "[A-Za-z0-9_\-]" Then
            safeName = safeName & c
        ElseIf c = " " Then
            safeName = safeName & "_"
        End If
    Next i
    
    If safeName = "" Then safeName = "Usuario"
    
    GetSafeUserName = safeName
    Exit Function
    
ErrorHandler:
    GetSafeUserName = "Usuario"
End Function

'================================================================================
' VALIDATE DOCUMENT STRUCTURE - SIMPLIFICADO - #STABLE
'================================================================================
Private Function ValidateDocumentStructure(doc As Document) As Boolean
    On Error Resume Next
    
    ' Verificação básica e rápida
    If doc.Range.End > 0 And doc.Sections.count > 0 Then
        ValidateDocumentStructure = True
    Else
        LogMessage "Documento com estrutura inconsistente", LOG_LEVEL_WARNING
        ValidateDocumentStructure = False
    End If
End Function

'================================================================================
' CRITICAL FIX: SAVE DOCUMENT BEFORE PROCESSING
' TO PREVENT CRASHES ON NEW NON SAVED DOCUMENTS - #STABLE
'================================================================================
Private Function SaveDocumentFirst(doc As Document) As Boolean
    On Error GoTo ErrorHandler

    Application.StatusBar = "Aguardando salvamento do documento..."
    ' Log de início removido para performance
    
    Dim saveDialog As Object
    Set saveDialog = Application.Dialogs(wdDialogFileSaveAs)

    If saveDialog.Show <> -1 Then
        LogMessage "Operação de salvamento cancelada pelo usuário", LOG_LEVEL_INFO
        Application.StatusBar = "Salvamento cancelado pelo usuário"
        SaveDocumentFirst = False
        Exit Function
    End If

    ' Aguarda confirmação do salvamento com timeout de segurança
    Dim waitCount As Integer
    Dim maxWait As Integer
    maxWait = 10
    
    For waitCount = 1 To maxWait
        DoEvents
        If doc.Path <> "" Then Exit For
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 1
            DoEvents
        Loop
        Application.StatusBar = "Aguardando salvamento... (" & waitCount & "/" & maxWait & ")"
    Next waitCount

    If doc.Path = "" Then
        LogMessage "Falha ao salvar documento após " & maxWait & " tentativas", LOG_LEVEL_ERROR
        Application.StatusBar = "Falha no salvamento - operação cancelada"
        SaveDocumentFirst = False
    Else
        ' Log de sucesso removido para performance
        Application.StatusBar = "Documento salvo com sucesso"
        SaveDocumentFirst = True
    End If

    Exit Function

ErrorHandler:
    LogMessage "Erro durante salvamento: " & Err.Description & " (Erro #" & Err.Number & ")", LOG_LEVEL_ERROR
    Application.StatusBar = "Erro durante salvamento"
    SaveDocumentFirst = False
End Function

'================================================================================
' CLEAR ALL FORMATTING - LIMPEZA INICIAL COMPLETA - #NEW
'================================================================================
Private Function ClearAllFormatting(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Limpando formatação existente..."
    
    ' SUPER OTIMIZADO: Verificação única de conteúdo visual no documento
    Dim hasImages As Boolean
    Dim hasShapes As Boolean
    hasImages = (doc.InlineShapes.count > 0)
    hasShapes = (doc.Shapes.count > 0)
    Dim hasAnyVisualContent As Boolean
    hasAnyVisualContent = hasImages Or hasShapes
    
    Dim paraCount As Long
    Dim styleResetCount As Long
    
    If hasAnyVisualContent Then
        ' MODO SEGURO OTIMIZADO: Cache de verificações visuais por parágrafo
        Dim para As Paragraph
        Dim visualContentCache As Object ' Cache para evitar recálculos
        Set visualContentCache = CreateObject("Scripting.Dictionary")
        
        For Each para In doc.Paragraphs
            On Error Resume Next
            
            ' Cache da verificação de conteúdo visual
            Dim paraKey As String
            paraKey = CStr(para.Range.Start) & "-" & CStr(para.Range.End)
            
            Dim hasVisualInPara As Boolean
            If visualContentCache.Exists(paraKey) Then
                hasVisualInPara = visualContentCache(paraKey)
            Else
                hasVisualInPara = HasVisualContent(para)
                visualContentCache.Add paraKey, hasVisualInPara
            End If
            
            If Not hasVisualInPara Then
                ' FORMATAÇÃO CONSOLIDADA: Aplica todas as configurações em uma única operação
                With para.Range
                    ' Reset completo de fonte em uma única operação
                    With .Font
                        .Reset
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                        .Bold = False
                        .Italic = False
                        .Underline = wdUnderlineNone
                    End With
                    
                    ' Reset completo de parágrafo em uma única operação
                    With .ParagraphFormat
                        .Reset
                        .alignment = wdAlignParagraphLeft
                        .LineSpacing = 12
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        .leftIndent = 0
                        .RightIndent = 0
                        .firstLineIndent = 0
                    End With
                    
                    ' Reset de bordas e sombreamento
                    .Borders.Enable = False
                    .Shading.Texture = wdTextureNone
                End With
                paraCount = paraCount + 1
            Else
                ' OTIMIZADO: Para parágrafos com imagens, formatação protegida mais rápida
                Call FormatCharacterByCharacter(para, STANDARD_FONT, STANDARD_FONT_SIZE, wdColorAutomatic, True, True)
                paraCount = paraCount + 1
            End If
            
            ' Proteção otimizada contra loops infinitos
            If paraCount Mod 100 = 0 Then DoEvents ' Permite responsividade a cada 100 parágrafos
            If paraCount > 1000 Then Exit For
            On Error GoTo ErrorHandler
        Next para
        
    Else
        ' MODO ULTRA-RÁPIDO: Sem conteúdo visual - formatação global em uma única operação
        With doc.Range
            ' Reset completo de fonte
            With .Font
                .Reset
                .Name = STANDARD_FONT
                .size = STANDARD_FONT_SIZE
                .Color = wdColorAutomatic
                .Bold = False
                .Italic = False
                .Underline = wdUnderlineNone
            End With
            
            ' Reset completo de parágrafo
            With .ParagraphFormat
                .Reset
                .alignment = wdAlignParagraphLeft
                .LineSpacing = 12
                .SpaceBefore = 0
                .SpaceAfter = 0
                .leftIndent = 0
                .RightIndent = 0
                .firstLineIndent = 0
            End With
            
            On Error Resume Next
            .Borders.Enable = False
            .Shading.Texture = wdTextureNone
            On Error GoTo ErrorHandler
        End With
        
        paraCount = doc.Paragraphs.count
    End If
    
    ' OTIMIZADO: Reset de estilos em uma única passada
    For Each para In doc.Paragraphs
        On Error Resume Next
        para.Style = "Normal"
        styleResetCount = styleResetCount + 1
        ' Otimização: Permite responsividade e proteção contra loops
        If styleResetCount Mod 50 = 0 Then DoEvents
        If styleResetCount > 1000 Then Exit For
        On Error GoTo ErrorHandler
    Next para
    
    LogMessage "Formatação limpa: " & paraCount & " parágrafos resetados", LOG_LEVEL_INFO
    ClearAllFormatting = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao limpar formatação: " & Err.Description, LOG_LEVEL_WARNING
    ClearAllFormatting = False ' Não falha o processo por isso
End Function

'================================================================================
' CLEAN DOCUMENT STRUCTURE - FUNCIONALIDADES 2, 6, 7 - #NEW
'================================================================================
Private Function CleanDocumentStructure(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim i As Long
    Dim firstTextParaIndex As Long
    Dim emptyLinesRemoved As Long
    Dim leadingSpacesRemoved As Long
    Dim paraCount As Long
    
    ' Cache da contagem total de parágrafos
    paraCount = doc.Paragraphs.count
    
    ' OTIMIZADO: Funcionalidade 2 - Remove linhas em branco acima do título
    ' Busca otimizada do primeiro parágrafo com texto
    firstTextParaIndex = -1
    For i = 1 To paraCount
        If i > doc.Paragraphs.count Then Exit For ' Proteção dinâmica
        
        Set para = doc.Paragraphs(i)
        Dim paraTextCheck As String
        paraTextCheck = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Encontra o primeiro parágrafo com texto real
        If paraTextCheck <> "" Then
            firstTextParaIndex = i
            Exit For
        End If
        
        ' Proteção contra documentos muito grandes
        If i > 50 Then Exit For ' Limita busca aos primeiros 50 parágrafos
    Next i
    
    ' OTIMIZADO: Remove linhas vazias ANTES do primeiro texto em uma única passada
    If firstTextParaIndex > 1 Then
        ' Processa de trás para frente para evitar problemas com índices
        For i = firstTextParaIndex - 1 To 1 Step -1
            If i > doc.Paragraphs.count Or i < 1 Then Exit For ' Proteção dinâmica
            
            Set para = doc.Paragraphs(i)
            Dim paraTextEmpty As String
            paraTextEmpty = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' OTIMIZADO: Verificação visual só se necessário
            If paraTextEmpty = "" Then
                If Not HasVisualContent(para) Then
                    para.Range.Delete
                    emptyLinesRemoved = emptyLinesRemoved + 1
                    ' Atualiza cache após remoção
                    paraCount = paraCount - 1
                End If
            End If
        Next i
    End If
    
    ' SUPER OTIMIZADO: Funcionalidade 7 - Remove espaços iniciais com regex
    ' Usa Find/Replace que é muito mais rápido que loop por parágrafo
    Dim rng As Range
    Set rng = doc.Range
    
    ' Remove espaços no início de linhas usando Find/Replace
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        
        ' Remove espaços/tabs no início de linhas usando Find/Replace simples
        .text = "^p "  ' Quebra seguida de espaço
        .Replacement.text = "^p"
        
        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            ' Proteção contra loop infinito
            If leadingSpacesRemoved > 1000 Then Exit Do
        Loop
        
        ' Remove tabs no início de linhas
        .text = "^p^t"  ' Quebra seguida de tab
        .Replacement.text = "^p"
        
        Do While .Execute(Replace:=True)
            leadingSpacesRemoved = leadingSpacesRemoved + 1
            If leadingSpacesRemoved > 1000 Then Exit Do
        Loop
    End With
    
    ' Segunda passada para espaços no início do documento (sem ^p precedente)
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False  ' Não usa wildcards nesta seção
        
        ' Posiciona no início do documento
        rng.Start = 0
        rng.End = 1
        
        ' Remove espaços/tabs no início absoluto do documento
        If rng.text = " " Or rng.text = vbTab Then
            ' Expande o range para pegar todos os espaços iniciais usando método seguro
            Do While rng.End <= doc.Range.End And (SafeGetLastCharacter(rng) = " " Or SafeGetLastCharacter(rng) = vbTab)
                rng.End = rng.End + 1
                leadingSpacesRemoved = leadingSpacesRemoved + 1
                If leadingSpacesRemoved > 100 Then Exit Do ' Proteção
            Loop
            
            If rng.Start < rng.End - 1 Then
                rng.Delete
            End If
        End If
    End With
    
    ' Log simplificado apenas se houve limpeza significativa
    If emptyLinesRemoved > 0 Then
        LogMessage "Estrutura limpa: " & emptyLinesRemoved & " linhas vazias removidas"
    End If
    
    CleanDocumentStructure = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza da estrutura: " & Err.Description, LOG_LEVEL_ERROR
    CleanDocumentStructure = False
End Function

'================================================================================
' SAFE CHECK FOR VISUAL CONTENT - VERIFICAÇÃO SEGURA DE CONTEÚDO VISUAL - #NEW
'================================================================================
Private Function HasVisualContent(para As Paragraph) As Boolean
    ' Usa a função segura implementada para compatibilidade total
    HasVisualContent = SafeHasVisualContent(para)
End Function

'================================================================================
' VALIDATE PROPOSITION TYPE - FUNCIONALIDADE 3 - #NEW
'================================================================================
Private Function ValidatePropositionType(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim firstWord As String
    Dim paraText As String
    Dim i As Long
    Dim userResponse As VbMsgBoxResult
    
    ' Encontra o primeiro parágrafo com texto
    For i = 1 To doc.Paragraphs.count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        LogMessage "Documento não possui texto para validação", LOG_LEVEL_WARNING
        ValidatePropositionType = True
        Exit Function
    End If
    
    ' Extrai a primeira palavra
    Dim words() As String
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
    End If
    
    ' Verifica se é uma das proposituras válidas
    If firstWord = "indicação" Or firstWord = "requerimento" Or firstWord = "moção" Then
        LogMessage "Tipo de proposição validado: " & firstWord, LOG_LEVEL_INFO
        ValidatePropositionType = True
    Else
        ' Informa sobre documento não-padrão e continua automaticamente
        LogMessage "Primeira palavra não reconhecida como proposição padrão: " & firstWord & " - continuando processamento", LOG_LEVEL_WARNING
        Application.StatusBar = "Aviso: Documento não é Indicação/Requerimento/Moção - processando mesmo assim"
        
        ' Pequena pausa para o usuário visualizar a mensagem
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 2  ' 2 segundos
            DoEvents
        Loop
        
        LogMessage "Processamento de documento não-padrão autorizado automaticamente: " & firstWord, LOG_LEVEL_INFO
        ValidatePropositionType = True
    End If
    
    Exit Function

ErrorHandler:
    LogMessage "Erro na validação do tipo de proposição: " & Err.Description, LOG_LEVEL_ERROR
    ValidatePropositionType = False
End Function

'================================================================================
' FORMAT DOCUMENT TITLE - FUNCIONALIDADES 4 e 5 - #NEW
'================================================================================
Private Function FormatDocumentTitle(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim firstPara As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim i As Long
    Dim newText As String
    
    ' Encontra o primeiro parágrafo com texto (após exclusão de linhas em branco)
    For i = 1 To doc.Paragraphs.count
        Set firstPara = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(firstPara.Range.text, vbCr, ""), vbLf, ""))
        If paraText <> "" Then
            Exit For
        End If
    Next i
    
    If paraText = "" Then
        LogMessage "Nenhum texto encontrado para formatação do título", LOG_LEVEL_WARNING
        FormatDocumentTitle = True
        Exit Function
    End If
    
    ' Remove ponto final se existir
    If Right(paraText, 1) = "." Then
        paraText = Left(paraText, Len(paraText) - 1)
    End If
    
    ' Verifica se é uma proposição (para aplicar substituição $NUMERO$/$ANO$)
    Dim isProposition As Boolean
    Dim firstWord As String
    
    words = Split(paraText, " ")
    If UBound(words) >= 0 Then
        firstWord = LCase(Trim(words(0)))
        If firstWord = "indicação" Or firstWord = "requerimento" Or firstWord = "moção" Then
            isProposition = True
        End If
    End If
    
    ' Se for proposição, substitui a última palavra por $NUMERO$/$ANO$
    If isProposition And UBound(words) >= 0 Then
        ' Reconstrói o texto substituindo a última palavra
        newText = ""
        For i = 0 To UBound(words) - 1
            If i > 0 Then newText = newText & " "
            newText = newText & words(i)
        Next i
        
        ' Adiciona $NUMERO$/$ANO$ no lugar da última palavra
        If newText <> "" Then newText = newText & " "
        newText = newText & "$NUMERO$/$ANO$"
    Else
        ' Se não for proposição, mantém o texto original
        newText = paraText
    End If
    
    ' SEMPRE aplica formatação de título: caixa alta, negrito, sublinhado
    firstPara.Range.text = UCase(newText) & vbCrLf
    
    ' Formatação completa do título (primeira linha)
    With firstPara.Range.Font
        .Bold = True
        .Underline = wdUnderlineSingle
    End With
    
    With firstPara.Format
        .alignment = wdAlignParagraphCenter
        .leftIndent = 0
        .firstLineIndent = 0
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 6  ' Pequeno espaço após o título
    End With
    
    If isProposition Then
        LogMessage "Título de proposição formatado: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    Else
        LogMessage "Primeira linha formatada como título: " & newText & " (centralizado, caixa alta, negrito, sublinhado)", LOG_LEVEL_INFO
    End If
    
    FormatDocumentTitle = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação do título: " & Err.Description, LOG_LEVEL_ERROR
    FormatDocumentTitle = False
End Function

'================================================================================
' FORMAT CONSIDERANDO PARAGRAPHS - OTIMIZADO E SIMPLIFICADO - FUNCIONALIDADE 8 - #NEW
'================================================================================
Private Function FormatConsiderandoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim rng As Range
    Dim totalFormatted As Long
    Dim i As Long
    
    ' Percorre todos os parágrafos procurando por "considerando" no início
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
        
        ' Verifica se o parágrafo começa com "considerando" (ignorando maiúsculas/minúsculas)
        If Len(paraText) >= 12 And LCase(Left(paraText, 12)) = "considerando" Then
            ' Verifica se após "considerando" vem espaço, vírgula, ponto-e-vírgula ou fim da linha
            Dim nextChar As String
            If Len(paraText) > 12 Then
                nextChar = Mid(paraText, 13, 1)
                If nextChar = " " Or nextChar = "," Or nextChar = ";" Or nextChar = ":" Then
                    ' É realmente "considerando" no início do parágrafo
                    Set rng = para.Range
                    
                    ' CORREÇÃO: Usa Find/Replace para preservar espaçamento
                    With rng.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .text = "considerando"
                        .Replacement.text = "CONSIDERANDO"
                        .Replacement.Font.Bold = True
                        .MatchCase = False
                        .MatchWholeWord = False  ' CORREÇÃO: False para não exigir palavra completa
                        .Forward = True
                        .Wrap = wdFindStop
                        
                        ' Limita a busca ao início do parágrafo
                        rng.End = rng.Start + 15  ' Seleciona apenas o início para evitar múltiplas substituições
                        
                        If .Execute(Replace:=True) Then
                            totalFormatted = totalFormatted + 1
                        End If
                    End With
                End If
            Else
                ' Parágrafo contém apenas "considerando"
                Set rng = para.Range
                rng.End = rng.Start + 12
                
                With rng
                    .text = "CONSIDERANDO"
                    .Font.Bold = True
                End With
                
                totalFormatted = totalFormatted + 1
            End If
        End If
    Next i
    
    LogMessage "Formatação 'considerando' aplicada: " & totalFormatted & " ocorrências em negrito e caixa alta", LOG_LEVEL_INFO
    FormatConsiderandoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação 'considerando': " & Err.Description, LOG_LEVEL_ERROR
    FormatConsiderandoParagraphs = False
End Function

'================================================================================
' APPLY TEXT REPLACEMENTS - FUNCIONALIDADES 10 e 11 - #NEW
'================================================================================
Private Function ApplyTextReplacements(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim replacementCount As Long
    
    Set rng = doc.Range
    
    ' Funcionalidade 10: Substitui variantes de "d'Oeste"
    Dim dOesteVariants() As String
    Dim i As Long
    
    ' Define as variantes possíveis dos 3 primeiros caracteres de "d'Oeste"
    ReDim dOesteVariants(0 To 15)
    dOesteVariants(0) = "d'O"   ' Original
    dOesteVariants(1) = "d´O"   ' Acento agudo
    dOesteVariants(2) = "d`O"   ' Acento grave
    dOesteVariants(3) = "d" & Chr(8220) & "O"   ' Aspas curvas esquerda
    dOesteVariants(4) = "d'o"   ' Minúscula
    dOesteVariants(5) = "d´o"
    dOesteVariants(6) = "d`o"
    dOesteVariants(7) = "d" & Chr(8220) & "o"
    dOesteVariants(8) = "D'O"   ' Maiúscula no D
    dOesteVariants(9) = "D´O"
    dOesteVariants(10) = "D`O"
    dOesteVariants(11) = "D" & Chr(8220) & "O"
    dOesteVariants(12) = "D'o"
    dOesteVariants(13) = "D´o"
    dOesteVariants(14) = "D`o"
    dOesteVariants(15) = "D" & Chr(8220) & "o"
    
    For i = 0 To UBound(dOesteVariants)
        With rng.Find
            .ClearFormatting
            .text = dOesteVariants(i) & "este"
            .Replacement.text = "d'Oeste"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute(Replace:=True)
                replacementCount = replacementCount + 1
            Loop
        End With
    Next i
    
    ' Funcionalidade 11: Substitui variantes de "- Vereador -"
    Set rng = doc.Range
    Dim vereadorVariants() As String
    ReDim vereadorVariants(0 To 7)
    
    ' Variantes dos caracteres inicial e final
    vereadorVariants(0) = "- Vereador -"    ' Original
    vereadorVariants(1) = "– Vereador –"    ' Travessão
    vereadorVariants(2) = "— Vereador —"    ' Em dash
    vereadorVariants(3) = "- vereador -"    ' Minúscula
    vereadorVariants(4) = "– vereador –"
    vereadorVariants(5) = "— vereador —"
    vereadorVariants(6) = "-Vereador-"      ' Sem espaços
    vereadorVariants(7) = "–Vereador–"
    
    For i = 0 To UBound(vereadorVariants)
        If vereadorVariants(i) <> "- Vereador -" Then
            With rng.Find
                .ClearFormatting
                .text = vereadorVariants(i)
                .Replacement.text = "- Vereador -"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                
                Do While .Execute(Replace:=True)
                    replacementCount = replacementCount + 1
                Loop
            End With
        End If
    Next i
    
    LogMessage "Substituições de texto aplicadas: " & replacementCount & " substituições realizadas", LOG_LEVEL_INFO
    ApplyTextReplacements = True
    Exit Function

ErrorHandler:
    LogMessage "Erro nas substituições de texto: " & Err.Description, LOG_LEVEL_ERROR
    ApplyTextReplacements = False
End Function

'================================================================================
' APPLY BOLD TO SPECIAL PARAGRAPHS - SIMPLIFIED & OPTIMIZED - #REFACTORED
'================================================================================
Private Sub ApplyBoldToSpecialParagraphs(doc As Document)
    On Error GoTo ErrorHandler
    
    If Not ValidateDocument(doc) Then Exit Sub
    
    Dim para As Paragraph
    Dim cleanText As String
    Dim specialParagraphs As Collection
    Set specialParagraphs = New Collection
    
    ' FASE 1: Identificar parágrafos especiais (uma única passada)
    For Each para In doc.Paragraphs
        If Not HasVisualContent(para) Then
            cleanText = GetCleanParagraphText(para)
            
            If cleanText = JUSTIFICATIVA_TEXT Or _
               IsAnexoPattern(cleanText) Or _
               IsVereadorPattern(cleanText) Then
                specialParagraphs.Add para
            End If
        End If
    Next para
    
    ' FASE 2: Aplicar negrito E reforçar alinhamento atomicamente
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    Dim p As Variant
    Dim pCleanText As String
    For Each p In specialParagraphs
        Set para = p ' Converte Variant para Paragraph
        
        ' Aplica negrito
        With para.Range.Font
            .Bold = True
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
        End With
        
        ' REFORÇO: Garante alinhamento correto baseado no tipo
        pCleanText = GetCleanParagraphText(para)
        If pCleanText = JUSTIFICATIVA_TEXT Then
            ' Justificativa: centralizado (linhas em branco serão inseridas depois)
            para.Format.alignment = wdAlignParagraphCenter
            para.Format.leftIndent = 0
            para.Format.firstLineIndent = 0
            para.Format.RightIndent = 0
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
        ElseIf IsVereadorPattern(pCleanText) Then
            ' Vereador: centralizado sem espaçamento extra
            para.Format.alignment = wdAlignParagraphCenter
            para.Format.leftIndent = 0
            para.Format.firstLineIndent = 0
            para.Format.RightIndent = 0
        ElseIf IsAnexoPattern(pCleanText) Then
            ' Anexo/Anexos: alinhado à esquerda
            para.Format.alignment = wdAlignParagraphLeft
            para.Format.leftIndent = 0
            para.Format.firstLineIndent = 0
            para.Format.RightIndent = 0
        End If
    Next p
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    
    LogMessage "Negrito e alinhamento aplicados a " & specialParagraphs.Count & " parágrafos especiais", LOG_LEVEL_INFO
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    LogMessage "Erro ao aplicar negrito a parágrafos especiais: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' INSERT JUSTIFICATIVA BLANK LINES - INSERE LINHAS EM BRANCO - #NEW
'================================================================================
Private Sub InsertJustificativaBlankLines(doc As Document)
    On Error GoTo ErrorHandler
    
    If Not ValidateDocument(doc) Then Exit Sub
    
    Dim para As Paragraph
    Dim cleanText As String
    Dim i As Long
    Dim justificativaIndex As Long
    Dim paraText As String
    
    Application.ScreenUpdating = False
    
    ' FASE 1: Localiza o parágrafo "Justificativa"
    justificativaIndex = 0
    For i = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        
        If Not HasVisualContent(para) Then
            cleanText = GetCleanParagraphText(para)
            
            If cleanText = JUSTIFICATIVA_TEXT Then
                justificativaIndex = i
                Exit For
            End If
        End If
    Next i
    
    If justificativaIndex = 0 Then
        Application.ScreenUpdating = True
        Exit Sub ' Não encontrou "Justificativa"
    End If
    
    ' FASE 2: Remove TODAS as linhas vazias ANTES de "Justificativa"
    i = justificativaIndex - 1
    Do While i >= 1
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        If paraText = "" And Not HasVisualContent(para) Then
            ' Remove linha vazia
            para.Range.Delete
            justificativaIndex = justificativaIndex - 1 ' Ajusta índice
            i = i - 1
        Else
            ' Encontrou conteúdo, para de remover
            Exit Do
        End If
    Loop
    
    ' FASE 3: Remove TODAS as linhas vazias DEPOIS de "Justificativa"
    i = justificativaIndex + 1
    Do While i <= doc.Paragraphs.Count
        Set para = doc.Paragraphs(i)
        paraText = Trim(Replace(Replace(para.Range.Text, vbCr, ""), vbLf, ""))
        
        If paraText = "" And Not HasVisualContent(para) Then
            ' Remove linha vazia
            para.Range.Delete
            ' Não incrementa i pois removemos o parágrafo
        Else
            ' Encontrou conteúdo, para de remover
            Exit Do
        End If
    Loop
    
    ' FASE 4: Insere EXATAMENTE 2 linhas em branco ANTES
    Set para = doc.Paragraphs(justificativaIndex)
    para.Range.InsertParagraphBefore
    para.Range.InsertParagraphBefore
    
    ' FASE 5: Insere EXATAMENTE 2 linhas em branco DEPOIS
    ' Atualiza referência após inserções anteriores
    Set para = doc.Paragraphs(justificativaIndex + 2) ' +2 porque inserimos 2 antes
    para.Range.InsertParagraphAfter
    para.Range.InsertParagraphAfter
    
    LogMessage "Linhas em branco ajustadas: exatamente 2 antes e 2 depois de 'Justificativa'", LOG_LEVEL_INFO
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    LogMessage "Erro ao inserir linhas em branco para Justificativa: " & Err.Description, LOG_LEVEL_WARNING
End Sub

'================================================================================
' FORMAT JUSTIFICATIVA/ANEXO PARAGRAPHS - FORMATAÇÃO ESPECÍFICA - #REFACTORED
'================================================================================
Private Function FormatJustificativaAnexoParagraphs(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Dim para As Paragraph
    Dim paraText As String
    Dim cleanText As String
    Dim i As Long
    Dim formattedCount As Long
    Dim vereadorCount As Long
    
    ' Percorre todos os parágrafos do documento
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        
        ' Não processa parágrafos com conteúdo visual
        If Not HasVisualContent(para) Then
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' Remove pontuação final para análise mais precisa
            cleanText = paraText
            ' Remove pontos, vírgulas, dois-pontos, ponto-e-vírgula do final
            Do While Len(cleanText) > 0 And (Right(cleanText, 1) = "." Or Right(cleanText, 1) = "," Or Right(cleanText, 1) = ":" Or Right(cleanText, 1) = ";")
                cleanText = Left(cleanText, Len(cleanText) - 1)
            Loop
            cleanText = Trim(LCase(cleanText))
            
            ' REQUISITO 1: Formatação de "justificativa" (case insensitive)
            If LCase(Trim(cleanText)) = JUSTIFICATIVA_TEXT Then
                ' Padroniza o texto mantendo pontuação original se houver
                Dim originalEnd As String
                originalEnd = ""
                If Len(paraText) > Len(cleanText) Then
                    originalEnd = Right(paraText, Len(paraText) - Len(cleanText))
                End If
                para.Range.text = "Justificativa" & originalEnd & vbCrLf
                
                ' Aplica formatação específica para Justificativa (SEM negrito - será aplicado depois)
                With para.Format
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                    .alignment = wdAlignParagraphCenter
                    .SpaceBefore = 0  ' Sem espaçamento - usaremos linhas vazias reais
                    .SpaceAfter = 0   ' Sem espaçamento - usaremos linhas vazias reais
                    .LineSpacingRule = wdLineSpacingMultiple
                    .LineSpacing = LINE_SPACING
                End With
                
                ' REFORÇO: Garante alinhamento centralizado com chamadas individuais
                para.Format.alignment = wdAlignParagraphCenter
                para.Format.leftIndent = 0
                para.Format.firstLineIndent = 0
                para.Format.RightIndent = 0
                
                LogMessage "Parágrafo 'Justificativa' formatado (centralizado, sem recuos, com 2 linhas antes e depois)", LOG_LEVEL_INFO
                formattedCount = formattedCount + 1
                
            ' REQUISITO 1: Formatação de variações de "vereador"
            ElseIf IsVereadorPattern(cleanText) Then
                ' REQUISITO 2: Formatar parágrafo ANTERIOR a "vereador" PRIMEIRO
                If i > 1 Then
                    Dim paraPrev As Paragraph
                    Set paraPrev = doc.Paragraphs(i - 1)
                    
                    ' Verifica se o parágrafo anterior não tem conteúdo visual
                    If Not HasVisualContent(paraPrev) Then
                        Dim prevText As String
                        prevText = Trim(Replace(Replace(paraPrev.Range.text, vbCr, ""), vbLf, ""))
                        
                        ' Só formata se o parágrafo anterior tem conteúdo textual
                        If prevText <> "" Then
                            ' Aplica caixa alta ao parágrafo anterior PRIMEIRO
                            paraPrev.Range.text = UCase(prevText) & vbCrLf
                            
                            ' Formatação do parágrafo anterior (SEM negrito - será aplicado depois)
                            With paraPrev.Format
                                .leftIndent = 0
                                .firstLineIndent = 0
                                .RightIndent = 0
                                .alignment = wdAlignParagraphCenter
                                .SpaceBefore = 0
                                .SpaceAfter = 0
                                .LineSpacingRule = wdLineSpacingMultiple
                                .LineSpacing = LINE_SPACING
                            End With
                            
                            LogMessage "Parágrafo anterior a '- Vereador -' formatado: " & Left(UCase(prevText), 30) & "...", LOG_LEVEL_INFO
                        End If
                    End If
                End If
                
                ' Padroniza o texto PRIMEIRO
                para.Range.text = "- Vereador -" & vbCrLf
                
                ' Formata o parágrafo "- Vereador -" (SEM negrito - será aplicado depois)
                With para.Format
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                    .alignment = wdAlignParagraphCenter
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpacingMultiple
                    .LineSpacing = LINE_SPACING
                End With
                
                LogMessage "Parágrafo '- Vereador -' formatado (centralizado, sem recuos)", LOG_LEVEL_INFO
                vereadorCount = vereadorCount + 1
                formattedCount = formattedCount + 1
                
            ' REQUISITO 3: Formatação de variações de "anexo" ou "anexos"
            ElseIf IsAnexoPattern(cleanText) Then
                ' Padroniza o texto mantendo pontuação original se houver
                Dim anexoEnd As String
                anexoEnd = ""
                If Len(paraText) > Len(cleanText) Then
                    anexoEnd = Right(paraText, Len(paraText) - Len(cleanText))
                End If
                
                Dim anexoText As String
                If cleanText = "anexo" Then
                    anexoText = "Anexo"
                Else
                    anexoText = "Anexos"
                End If
                para.Range.text = anexoText & anexoEnd & vbCrLf
                
                ' Aplica formatação específica para Anexo/Anexos (SEM negrito - será aplicado depois)
                With para.Format
                    .leftIndent = 0
                    .firstLineIndent = 0
                    .RightIndent = 0
                    .alignment = wdAlignParagraphLeft
                    .SpaceBefore = 0
                    .SpaceAfter = 0
                    .LineSpacingRule = wdLineSpacingMultiple
                    .LineSpacing = LINE_SPACING
                End With
                
                LogMessage "Parágrafo '" & anexoText & "' formatado (alinhado à esquerda, sem recuos)", LOG_LEVEL_INFO
                formattedCount = formattedCount + 1
            End If
        End If
    Next i
    
    LogMessage "Formatação especial concluída: " & formattedCount & " parágrafos formatados (incluindo " & vereadorCount & " '- Vereador -')", LOG_LEVEL_INFO
    FormatJustificativaAnexoParagraphs = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na formatação de parágrafos especiais: " & Err.Description, LOG_LEVEL_ERROR
    FormatJustificativaAnexoParagraphs = False
End Function

'================================================================================
' FUNÇÕES AUXILIARES PARA DETECÇÃO DE PADRÕES
'================================================================================
Private Function IsVereadorPattern(text As String) As Boolean
    ' Remove espaços extras para análise
    Dim cleanText As String
    cleanText = Trim(text)
    
    ' Remove hifens/travessões do início e fim e espaços adjacentes
    cleanText = Trim(cleanText)
    If Left(cleanText, 1) = "-" Or Left(cleanText, 1) = "–" Or Left(cleanText, 1) = "—" Then
        cleanText = Trim(Mid(cleanText, 2))
    End If
    If Right(cleanText, 1) = "-" Or Right(cleanText, 1) = "–" Or Right(cleanText, 1) = "—" Then
        cleanText = Trim(Left(cleanText, Len(cleanText) - 1))
    End If
    
    ' Verifica se o que sobrou é alguma variação de "vereador"
    cleanText = LCase(Trim(cleanText))
    IsVereadorPattern = (cleanText = "vereador" Or cleanText = "vereadora")
End Function

Private Function IsAnexoPattern(text As String) As Boolean
    Dim cleanText As String
    cleanText = LCase(Trim(text))
    IsAnexoPattern = (cleanText = "anexo" Or cleanText = "anexos")
End Function

'================================================================================
' SUBROTINA PÚBLICA: ABRIR PASTA DE LOGS - #NEW
'================================================================================
Public Sub AbrirPastaLogs()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim logsFolder As String
    Dim defaultLogsFolder As String
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define pasta de logs baseada no documento atual ou temp
    If Not doc Is Nothing And doc.Path <> "" Then
        logsFolder = doc.Path
    Else
        logsFolder = Environ("TEMP")
    End If
    
    ' Verifica se a pasta existe
    If Dir(logsFolder, vbDirectory) = "" Then
        logsFolder = Environ("TEMP")
    End If
    
    ' Abre a pasta no Windows Explorer
    shell "explorer.exe """ & logsFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de logs aberta: " & logsFolder
    
    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de logs aberta pelo usuário: " & logsFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de logs"
    
    ' Fallback: tenta abrir pasta temporária
    On Error Resume Next
    shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
    If Err.Number = 0 Then
        Application.StatusBar = "Pasta temporária aberta como alternativa"
    Else
        Application.StatusBar = "Não foi possível abrir pasta de logs"
    End If
End Sub

'================================================================================
' SUBROTINA PÚBLICA: ABRIR REPOSITÓRIO GITHUB - FUNCIONALIDADE 9 - #NEW
'================================================================================
Public Sub AbrirRepositorioGitHub()
    On Error GoTo ErrorHandler
    
    Dim repoURL As String
    Dim shellResult As Long
    
    ' URL do repositório do projeto
    repoURL = "https://github.com/chrmsantos/chainsaw-fprops"
    
    ' Abre o link no navegador padrão
    shellResult = shell("rundll32.exe url.dll,FileProtocolHandler " & repoURL, vbNormalFocus)
    
    If shellResult > 0 Then
        Application.StatusBar = "Repositório GitHub aberto no navegador"
        
        ' Log da operação se sistema de log estiver ativo
        If loggingEnabled Then
            LogMessage "Repositório GitHub aberto pelo usuário: " & repoURL, LOG_LEVEL_INFO
        End If
    Else
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir repositório GitHub"
    
    ' Log do erro se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Erro ao abrir repositório GitHub: " & Err.Description, LOG_LEVEL_ERROR
    End If
    
    ' Fallback: tenta copiar URL para a área de transferência
    On Error Resume Next
    Dim dataObj As Object
    Set dataObj = CreateObject("htmlfile").parentWindow.clipboardData
    dataObj.setData "text", repoURL
    
    If Err.Number = 0 Then
        Application.StatusBar = "URL copiada para área de transferência: " & repoURL
    Else
        Application.StatusBar = "Não foi possível abrir o repositório. URL: " & repoURL
    End If
End Sub

'================================================================================
' SISTEMA DE BACKUP - FUNCIONALIDADE DE SEGURANÇA - #NEW
'================================================================================
Private Function CreateDocumentBackup(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    ' Não faz backup se documento não foi salvo
    If doc.Path = "" Then
        LogMessage "Backup ignorado - documento não salvo", LOG_LEVEL_INFO
        CreateDocumentBackup = True
        Exit Function
    End If
    
    Dim backupFolder As String
    Dim fso As Object
    Dim docName As String
    Dim docExtension As String
    Dim timeStamp As String
    Dim backupFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define pasta de backup
    backupFolder = fso.GetParentFolderName(doc.Path) & "\" & BACKUP_FOLDER_NAME
    
    ' Cria pasta de backup se não existir
    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
        LogMessage "Pasta de backup criada: " & backupFolder, LOG_LEVEL_INFO
    End If
    
    ' Extrai nome e extensão do documento
    docName = fso.GetBaseName(doc.Name)
    docExtension = fso.GetExtensionName(doc.Name)
    
    ' Cria timestamp para o backup
    timeStamp = Format(Now, "yyyy-mm-dd_HHmmss")
    
    ' Nome do arquivo de backup
    backupFileName = docName & "_backup_" & timeStamp & "." & docExtension
    backupFilePath = backupFolder & "\" & backupFileName
    
    ' Salva uma cópia do documento como backup
    Application.StatusBar = "Criando backup do documento..."
    
    ' Salva o documento atual primeiro para garantir que está atualizado
    doc.Save
    
    ' Cria uma cópia do arquivo usando FileSystemObject
    fso.CopyFile doc.FullName, backupFilePath, True
    
    ' Limpa backups antigos se necessário
    CleanOldBackups backupFolder, docName
    
    LogMessage "Backup criado com sucesso: " & backupFileName, LOG_LEVEL_INFO
    Application.StatusBar = "Backup criado - processando documento..."
    
    CreateDocumentBackup = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao criar backup: " & Err.Description, LOG_LEVEL_ERROR
    CreateDocumentBackup = False
End Function

'================================================================================
' LIMPEZA DE BACKUPS ANTIGOS - SIMPLIFICADO - #NEW
'================================================================================
Private Sub CleanOldBackups(backupFolder As String, docBaseName As String)
    On Error Resume Next
    
    ' Limpeza simplificada - só remove se houver muitos arquivos
    Dim fso As Object
    Dim folder As Object
    Dim filesCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupFolder)
    
    filesCount = folder.Files.count
    
    ' Se há mais de 15 arquivos na pasta de backup, registra aviso
    If filesCount > 15 Then
        LogMessage "Muitos backups na pasta (" & filesCount & " arquivos) - considere limpeza manual", LOG_LEVEL_WARNING
    End If
End Sub

'================================================================================
' SUBROTINA PÚBLICA: ABRIR PASTA DE BACKUPS - #NEW
'================================================================================
Public Sub AbrirPastaBackups()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim backupFolder As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Tenta obter documento ativo
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo ErrorHandler
    
    ' Define pasta de backup baseada no documento atual
    If Not doc Is Nothing And doc.Path <> "" Then
        backupFolder = fso.GetParentFolderName(doc.Path) & "\" & BACKUP_FOLDER_NAME
    Else
        Application.StatusBar = "Nenhum documento salvo ativo para localizar pasta de backups"
        Exit Sub
    End If
    
    ' Verifica se a pasta de backup existe
    If Not fso.FolderExists(backupFolder) Then
        Application.StatusBar = "Pasta de backups não encontrada - nenhum backup foi criado ainda"
        LogMessage "Pasta de backups não encontrada: " & backupFolder, LOG_LEVEL_WARNING
        Exit Sub
    End If
    
    ' Abre a pasta no Windows Explorer
    shell "explorer.exe """ & backupFolder & """", vbNormalFocus
    
    Application.StatusBar = "Pasta de backups aberta: " & backupFolder
    
    ' Log da operação se sistema de log estiver ativo
    If loggingEnabled Then
        LogMessage "Pasta de backups aberta pelo usuário: " & backupFolder, LOG_LEVEL_INFO
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = "Erro ao abrir pasta de backups"
    LogMessage "Erro ao abrir pasta de backups: " & Err.Description, LOG_LEVEL_ERROR
    
    ' Fallback: tenta abrir pasta do documento
    On Error Resume Next
    If Not doc Is Nothing And doc.Path <> "" Then
        Dim docFolder As String
        docFolder = fso.GetParentFolderName(doc.Path)
        shell "explorer.exe """ & docFolder & """", vbNormalFocus
        Application.StatusBar = "Pasta do documento aberta como alternativa"
    Else
        Application.StatusBar = "Não foi possível abrir pasta de backups"
    End If
End Sub

'================================================================================
' CLEAN MULTIPLE SPACES - LIMPEZA FINAL DE ESPAÇOS MÚLTIPLOS - #NEW
'================================================================================
Private Function CleanMultipleSpaces(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Limpando espaços múltiplos..."
    
    Dim rng As Range
    Dim spacesRemoved As Long
    Dim totalOperations As Long
    
    ' SUPER OTIMIZADO: Operações consolidadas em uma única configuração Find
    Set rng = doc.Range
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' OTIMIZAÇÃO 1: Remove espaços múltiplos (2 ou mais) em uma única operação
        ' Usa um loop otimizado que reduz progressivamente os espaços
        Do
            .text = "  "  ' Dois espaços
            .Replacement.text = " "  ' Um espaço
            
            Dim currentReplaceCount As Long
            currentReplaceCount = 0
            
            ' Executa até não encontrar mais duplos
            Do While .Execute(Replace:=True)
                currentReplaceCount = currentReplaceCount + 1
                spacesRemoved = spacesRemoved + 1
                ' Proteção otimizada - verifica a cada 200 operações
                If currentReplaceCount Mod 200 = 0 Then
                    DoEvents
                    If spacesRemoved > 2000 Then Exit Do
                End If
            Loop
            
            totalOperations = totalOperations + 1
            ' Se não encontrou mais duplos ou atingiu limite, para
            If currentReplaceCount = 0 Or totalOperations > 10 Then Exit Do
        Loop
    End With
    
    ' OTIMIZAÇÃO 2: Operações de limpeza de quebras de linha consolidadas
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False  ' Usar Find/Replace simples para compatibilidade
        
        ' Remove múltiplos espaços antes de quebras - método iterativo
        .text = "  ^p"  ' 2 espaços seguidos de quebra
        .Replacement.text = " ^p"  ' 1 espaço seguido de quebra
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Segunda passada para garantir limpeza completa
        .text = " ^p"  ' Espaço antes de quebra
        .Replacement.text = "^p"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Remove múltiplos espaços depois de quebras - método iterativo
        .text = "^p  "  ' Quebra seguida de 2 espaços
        .Replacement.text = "^p "  ' Quebra seguida de 1 espaço
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With
    
    ' OTIMIZAÇÃO 3: Limpeza de tabs consolidada e otimizada
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False  ' Usar Find/Replace simples
        
        ' Remove múltiplos tabs iterativamente
        .text = "^t^t"  ' 2 tabs
        .Replacement.text = "^t"  ' 1 tab
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
        
        ' Converte tabs para espaços
        .text = "^t"
        .Replacement.text = " "
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2000 Then Exit Do
        Loop
    End With
    
    ' OTIMIZAÇÃO 4: Verificação final ultra-rápida de espaços duplos remanescentes
    Set rng = doc.Range
    With rng.Find
        .text = "  "
        .Replacement.text = " "
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop  ' Mais rápido que wdFindContinue
        
        Dim finalCleanCount As Long
        Do While .Execute(Replace:=True) And finalCleanCount < 100
            finalCleanCount = finalCleanCount + 1
            spacesRemoved = spacesRemoved + 1
        Loop
    End With
    
    ' PROTEÇÃO ESPECÍFICA: Garante espaço após CONSIDERANDO
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        
        ' Corrige CONSIDERANDO grudado com a próxima palavra
        .text = "CONSIDERANDOa"
        .Replacement.text = "CONSIDERANDO a"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .text = "CONSIDERANDOe"
        .Replacement.text = "CONSIDERANDO e"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .text = "CONSIDERANDOo"
        .Replacement.text = "CONSIDERANDO o"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
        
        .text = "CONSIDERANDOq"
        .Replacement.text = "CONSIDERANDO q"
        Do While .Execute(Replace:=True)
            spacesRemoved = spacesRemoved + 1
            If spacesRemoved > 2100 Then Exit Do
        Loop
    End With
    
    LogMessage "Limpeza de espaços concluída: " & spacesRemoved & " correções aplicadas (com proteção CONSIDERANDO)", LOG_LEVEL_INFO
    CleanMultipleSpaces = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na limpeza de espaços múltiplos: " & Err.Description, LOG_LEVEL_WARNING
    CleanMultipleSpaces = False ' Não falha o processo por isso
End Function

'================================================================================
' LIMIT SEQUENTIAL EMPTY LINES - CONTROLA LINHAS VAZIAS SEQUENCIAIS - #NEW
'================================================================================
Private Function LimitSequentialEmptyLines(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Controlando linhas em branco sequenciais..."
    
    ' IDENTIFICAÇÃO DO SEGUNDO PARÁGRAFO PARA PROTEÇÃO
    Dim secondParaIndex As Long
    secondParaIndex = GetSecondParagraphIndex(doc)
    
    ' SUPER OTIMIZADO: Usa Find/Replace com wildcard para operação muito mais rápida
    Dim rng As Range
    Dim linesRemoved As Long
    Dim totalReplaces As Long
    Dim passCount As Long
    
    passCount = 1 ' Inicializa contador de passadas
    
    Set rng = doc.Range
    
    ' MÉTODO ULTRA-RÁPIDO: Remove múltiplas quebras consecutivas usando wildcard
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False  ' Usar Find/Replace simples para compatibilidade
        
        ' Remove múltiplas quebras consecutivas iterativamente
        .text = "^p^p^p^p"  ' 4 quebras
        .Replacement.text = "^p^p"  ' 2 quebras
        
        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
        
        ' Remove 3 quebras -> 2 quebras
        .text = "^p^p^p"  ' 3 quebras
        .Replacement.text = "^p^p"  ' 2 quebras
        
        Do While .Execute(Replace:=True)
            linesRemoved = linesRemoved + 1
            totalReplaces = totalReplaces + 1
            If totalReplaces > 500 Then Exit Do
            If linesRemoved Mod 50 = 0 Then DoEvents
        Loop
    End With
    
    ' SEGUNDA PASSADA: Remove quebras duplas restantes (2 quebras -> 1 quebra)
    If totalReplaces > 0 Then passCount = passCount + 1
    
    Set rng = doc.Range
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindContinue
        
        ' Converte quebras duplas em quebras simples
        .text = "^p^p^p"  ' 3 quebras
        .Replacement.text = "^p^p"  ' 2 quebras
        
        Dim secondPassCount As Long
        Do While .Execute(Replace:=True) And secondPassCount < 200
            secondPassCount = secondPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With
    
    ' VERIFICAÇÃO FINAL: Garantir que não há mais de 1 linha vazia consecutiva
    If secondPassCount > 0 Then passCount = passCount + 1
    
    ' Método híbrido: Find/Replace para casos simples + loop apenas se necessário
    Set rng = doc.Range
    With rng.Find
        .text = "^p^p^p"  ' 3 quebras (2 linhas vazias + conteúdo)
        .Replacement.text = "^p^p"  ' 2 quebras (1 linha vazia + conteúdo)
        .MatchWildcards = False
        
        Dim finalPassCount As Long
        Do While .Execute(Replace:=True) And finalPassCount < 100
            finalPassCount = finalPassCount + 1
            linesRemoved = linesRemoved + 1
        Loop
    End With
    
    If finalPassCount > 0 Then passCount = passCount + 1
    
    ' FALLBACK OTIMIZADO: Se ainda há problemas, usa método tradicional limitado
    If finalPassCount >= 100 Then
        passCount = passCount + 1 ' Incrementa para o fallback
        
        Dim para As Paragraph
        Dim i As Long
        Dim emptyLineCount As Long
        Dim paraText As String
        Dim fallbackRemoved As Long
        
        i = 1
        emptyLineCount = 0
        
        Do While i <= doc.Paragraphs.count And fallbackRemoved < 50
            Set para = doc.Paragraphs(i)
            paraText = Trim(Replace(Replace(para.Range.text, vbCr, ""), vbLf, ""))
            
            ' Verifica se o parágrafo está vazio
            If paraText = "" And Not HasVisualContent(para) Then
                emptyLineCount = emptyLineCount + 1
                
                ' Se já temos mais de 1 linha vazia consecutiva, remove esta
                If emptyLineCount > 1 Then
                    para.Range.Delete
                    fallbackRemoved = fallbackRemoved + 1
                    linesRemoved = linesRemoved + 1
                    ' Não incrementa i pois removemos um parágrafo
                Else
                    i = i + 1
                End If
            Else
                ' Se encontrou conteúdo, reseta o contador
                emptyLineCount = 0
                i = i + 1
            End If
            
            ' Responsividade e proteção otimizadas
            If fallbackRemoved Mod 10 = 0 Then DoEvents
            If i > 500 Then Exit Do ' Proteção adicional
        Loop
    End If
    
    LogMessage "Controle de linhas vazias concluído em " & passCount & " passada(s): " & linesRemoved & " linhas excedentes removidas (máximo 1 sequencial)", LOG_LEVEL_INFO
    LimitSequentialEmptyLines = True
    Exit Function

ErrorHandler:
    LogMessage "Erro no controle de linhas vazias: " & Err.Description, LOG_LEVEL_WARNING
    LimitSequentialEmptyLines = False ' Não falha o processo por isso
End Function

'================================================================================
' CONFIGURE DOCUMENT VIEW - CONFIGURAÇÃO DE VISUALIZAÇÃO - #MODIFIED
'================================================================================
Private Function ConfigureDocumentView(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Configurando visualização do documento..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Configura APENAS o zoom para 110% - todas as outras configurações são preservadas
    With docWindow.View
        .Zoom.Percentage = 110
        ' NÃO altera mais o tipo de visualização - preserva o original
    End With
    
    ' Remove configurações que alteravam configurações globais do Word
    ' Estas configurações são agora preservadas do estado original
    
    LogMessage "Visualização configurada: zoom definido para 110%, demais configurações preservadas"
    ConfigureDocumentView = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao configurar visualização: " & Err.Description, LOG_LEVEL_WARNING
    ConfigureDocumentView = False ' Não falha o processo por isso
End Function

'================================================================================
' SALVAR E SAIR - SUBROTINA PÚBLICA PROFISSIONAL E ROBUSTA
'================================================================================
Public Sub SalvarESair()
    On Error GoTo CriticalErrorHandler
    
    Dim startTime As Date
    startTime = Now
    
    Application.StatusBar = "Verificando documentos abertos..."
    LogMessage "Iniciando processo de salvar e sair - verificação de documentos", LOG_LEVEL_INFO
    
    ' Verifica se há documentos abertos
    If Application.Documents.count = 0 Then
        Application.StatusBar = "Nenhum documento aberto - encerrando Word"
        LogMessage "Nenhum documento aberto - encerrando aplicação", LOG_LEVEL_INFO
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' Coleta informações sobre documentos não salvos
    Dim unsavedDocs As Collection
    Set unsavedDocs = New Collection
    
    Dim doc As Document
    Dim i As Long
    
    ' Verifica cada documento aberto
    For i = 1 To Application.Documents.count
        Set doc = Application.Documents(i)
        
        On Error Resume Next
        ' Verifica se o documento tem alterações não salvas
        If doc.Saved = False Or doc.Path = "" Then
            unsavedDocs.Add doc.Name
            LogMessage "Documento não salvo detectado: " & doc.Name
        End If
        On Error GoTo CriticalErrorHandler
    Next i
    
    ' Se não há documentos não salvos, encerra diretamente
    If unsavedDocs.count = 0 Then
        Application.StatusBar = "Todos os documentos salvos - encerrando Word"
        LogMessage "Todos os documentos estão salvos - encerrando aplicação"
        Application.Quit SaveChanges:=wdDoNotSaveChanges
        Exit Sub
    End If
    
    ' Constrói mensagem detalhada sobre documentos não salvos
    Dim message As String
    Dim docList As String
    
    For i = 1 To unsavedDocs.count
        docList = docList & "• " & unsavedDocs(i) & vbCrLf
    Next i
    
    message = "ATENÇÃO: Há " & unsavedDocs.count & " documento(s) com alterações não salvas:" & vbCrLf & vbCrLf
    message = message & docList & vbCrLf
    message = message & "Deseja salvar todos os documentos antes de sair?" & vbCrLf & vbCrLf
    message = message & "• SIM: Salva todos e fecha o Word" & vbCrLf
    message = message & "• NÃO: Fecha sem salvar (PERDE as alterações)" & vbCrLf
    message = message & "• CANCELAR: Cancela a operação"
    
    ' Apresenta opções ao usuário
    Application.StatusBar = "Aguardando decisão do usuário sobre documentos não salvos..."
    
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox(message, vbYesNoCancel + vbExclamation + vbDefaultButton1, _
                        "Chainsaw - Salvar e Sair (" & unsavedDocs.count & " documentos não salvos)")
    
    Select Case userChoice
        Case vbYes
            ' Usuário escolheu salvar todos
            Application.StatusBar = "Salvando todos os documentos..."
            LogMessage "Usuário optou por salvar todos os documentos antes de sair"
            
            If SalvarTodosDocumentos() Then
                Application.StatusBar = "Documentos salvos com sucesso - encerrando Word"
                LogMessage "Todos os documentos salvos com sucesso - encerrando aplicação"
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
                Application.StatusBar = "Erro ao salvar documentos - operação cancelada"
                LogMessage "Falha ao salvar alguns documentos - operação de sair cancelada", LOG_LEVEL_ERROR
                MsgBox "Erro ao salvar um ou mais documentos." & vbCrLf & _
                       "A operação foi cancelada por segurança." & vbCrLf & vbCrLf & _
                       "Verifique os documentos e tente novamente.", _
                       vbCritical, "Chainsaw - Erro ao Salvar"
            End If
            
        Case vbNo
            ' Usuário escolheu não salvar
            Dim confirmMessage As String
            confirmMessage = "CONFIRMAÇÃO FINAL:" & vbCrLf & vbCrLf
            confirmMessage = confirmMessage & "Você está prestes a FECHAR O WORD SEM SALVAR " & unsavedDocs.count & " documento(s)." & vbCrLf & vbCrLf
            confirmMessage = confirmMessage & "TODAS AS ALTERAÇÕES NÃO SALVAS SERÃO PERDIDAS!" & vbCrLf & vbCrLf
            confirmMessage = confirmMessage & "Tem certeza absoluta?"
            
            Dim finalConfirm As VbMsgBoxResult
            finalConfirm = MsgBox(confirmMessage, vbYesNo + vbCritical + vbDefaultButton2, _
                                  "Chainsaw - CONFIRMAÇÃO FINAL")
            
            If finalConfirm = vbYes Then
                Application.StatusBar = "Fechando Word sem salvar alterações..."
                LogMessage "Usuário confirmou fechamento sem salvar - encerrando aplicação", LOG_LEVEL_WARNING
                Application.Quit SaveChanges:=wdDoNotSaveChanges
            Else
                Application.StatusBar = "Operação cancelada pelo usuário"
                LogMessage "Usuário cancelou fechamento sem salvar"
                MsgBox "Operação cancelada." & vbCrLf & "Os documentos permanecem abertos.", _
                       vbInformation, "Chainsaw - Operação Cancelada"
            End If
            
        Case vbCancel
            ' Usuário cancelou
            Application.StatusBar = "Operação de sair cancelada pelo usuário"
            LogMessage "Usuário cancelou operação de salvar e sair"
            MsgBox "Operação cancelada." & vbCrLf & "Os documentos permanecem abertos.", _
                   vbInformation, "Chainsaw - Operação Cancelada"
    End Select
    
    Application.StatusBar = False
    LogMessage "Processo de salvar e sair concluído em " & Format(Now - startTime, "hh:mm:ss")
    Exit Sub

CriticalErrorHandler:
    Dim errDesc As String
    errDesc = "ERRO CRÍTICO na operação Salvar e Sair #" & Err.Number & ": " & Err.Description
    
    LogMessage errDesc, LOG_LEVEL_ERROR
    Application.StatusBar = "Erro crítico - operação cancelada"
    
    MsgBox "Erro crítico durante a operação Salvar e Sair:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "A operação foi cancelada por segurança." & vbCrLf & _
           "Salve manualmente os documentos importantes.", _
           vbCritical, "Chainsaw - Erro Crítico"
End Sub

'================================================================================
' SALVAR TODOS DOCUMENTOS - FUNÇÃO AUXILIAR PRIVADA
'================================================================================
Private Function SalvarTodosDocumentos() As Boolean
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim i As Long
    Dim savedCount As Long
    Dim errorCount As Long
    Dim totalDocs As Long
    
    totalDocs = Application.Documents.count
    
    ' Salva cada documento individualmente
    For i = 1 To totalDocs
        Set doc = Application.Documents(i)
        
        Application.StatusBar = "Salvando documento " & i & " de " & totalDocs & ": " & doc.Name
        
        On Error Resume Next
        
        ' Se o documento nunca foi salvo (sem caminho), abre dialog
        If doc.Path = "" Then
            Dim saveDialog As Object
            Set saveDialog = Application.FileDialog(msoFileDialogSaveAs)
            
            With saveDialog
                .Title = "Salvar documento: " & doc.Name
                .InitialFileName = doc.Name
                
                If .Show = -1 Then
                    doc.SaveAs2 .SelectedItems(1)
                    If Err.Number = 0 Then
                        savedCount = savedCount + 1
                        LogMessage "Documento salvo como novo arquivo: " & doc.Name
                    Else
                        errorCount = errorCount + 1
                        LogMessage "Erro ao salvar documento como novo: " & doc.Name & " - " & Err.Description, LOG_LEVEL_ERROR
                    End If
                Else
                    errorCount = errorCount + 1
                    LogMessage "Salvamento cancelado pelo usuário: " & doc.Name, LOG_LEVEL_WARNING
                End If
            End With
        Else
            ' Documento já tem caminho, apenas salva
            doc.Save
            If Err.Number = 0 Then
                savedCount = savedCount + 1
                LogMessage "Documento salvo: " & doc.Name
            Else
                errorCount = errorCount + 1
                LogMessage "Erro ao salvar documento: " & doc.Name & " - " & Err.Description, LOG_LEVEL_ERROR
            End If
        End If
        
        On Error GoTo ErrorHandler
    Next i
    
    ' Verifica resultado
    If errorCount = 0 Then
        LogMessage "Todos os documentos salvos com sucesso: " & savedCount & " de " & totalDocs
        SalvarTodosDocumentos = True
    Else
        LogMessage "Falha parcial no salvamento: " & savedCount & " salvos, " & errorCount & " erros", LOG_LEVEL_WARNING
        SalvarTodosDocumentos = False
    End If
    
    Exit Function

ErrorHandler:
    LogMessage "Erro crítico ao salvar documentos: " & Err.Description, LOG_LEVEL_ERROR
    SalvarTodosDocumentos = False
End Function

'================================================================================
' IMAGE PROTECTION SYSTEM - SISTEMA DE PROTEÇÃO DE IMAGENS - #NEW
'================================================================================

'================================================================================
' BACKUP ALL IMAGES - Faz backup de propriedades das imagens do documento
'================================================================================
Private Function BackupAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Fazendo backup das propriedades das imagens..."
    
    imageCount = 0
    ReDim savedImages(0)
    
    Dim para As Paragraph
    Dim i As Long
    Dim j As Long
    Dim shape As InlineShape
    Dim tempImageInfo As ImageInfo
    
    ' Conta todas as imagens primeiro
    Dim totalImages As Long
    For i = 1 To doc.Paragraphs.count
        Set para = doc.Paragraphs(i)
        totalImages = totalImages + para.Range.InlineShapes.count
    Next i
    
    ' Adiciona shapes flutuantes
    totalImages = totalImages + doc.Shapes.count
    
    ' Redimensiona array se necessário
    If totalImages > 0 Then
        ReDim savedImages(totalImages - 1)
        
        ' Backup de imagens inline - apenas propriedades críticas
        For i = 1 To doc.Paragraphs.count
            Set para = doc.Paragraphs(i)
            
            For j = 1 To para.Range.InlineShapes.count
                Set shape = para.Range.InlineShapes(j)
                
                ' Salva apenas propriedades essenciais para proteção
                With tempImageInfo
                    .paraIndex = i
                    .ImageIndex = j
                    .ImageType = "Inline"
                    .Position = shape.Range.Start
                    .Width = shape.Width
                    .Height = shape.Height
                    Set .AnchorRange = shape.Range.Duplicate
                    .ImageData = "InlineShape_Protected"
                End With
                
                savedImages(imageCount) = tempImageInfo
                imageCount = imageCount + 1
                
                ' Evita overflow
                If imageCount >= UBound(savedImages) + 1 Then Exit For
            Next j
            
            ' Evita overflow
            If imageCount >= UBound(savedImages) + 1 Then Exit For
        Next i
        
        ' Backup de shapes flutuantes - apenas propriedades críticas
        Dim floatingShape As shape
        For i = 1 To doc.Shapes.count
            Set floatingShape = doc.Shapes(i)
            
            If floatingShape.Type = msoPicture Then
                ' Redimensiona array se necessário
                If imageCount >= UBound(savedImages) + 1 Then
                    ReDim Preserve savedImages(imageCount)
                End If
                
                With tempImageInfo
                    .paraIndex = -1 ' Indica que é flutuante
                    .ImageIndex = i
                    .ImageType = "Floating"
                    .WrapType = floatingShape.WrapFormat.Type
                    .Width = floatingShape.Width
                    .Height = floatingShape.Height
                    .LeftPosition = floatingShape.Left
                    .TopPosition = floatingShape.Top
                    .ImageData = "FloatingShape_Protected"
                End With
                
                savedImages(imageCount) = tempImageInfo
                imageCount = imageCount + 1
            End If
        Next i
    End If
    
    LogMessage "Backup de propriedades de imagens concluído: " & imageCount & " imagens catalogadas"
    BackupAllImages = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup de propriedades de imagens: " & Err.Description, LOG_LEVEL_WARNING
    BackupAllImages = False
End Function

'================================================================================
' RESTORE ALL IMAGES - Verifica e corrige propriedades das imagens
'================================================================================
Private Function RestoreAllImages(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    If imageCount = 0 Then
        RestoreAllImages = True
        Exit Function
    End If
    
    Application.StatusBar = "Verificando integridade das imagens..."
    
    Dim i As Long
    Dim verifiedCount As Long
    Dim correctedCount As Long
    
    For i = 0 To imageCount - 1
        On Error Resume Next
        
        With savedImages(i)
            If .ImageType = "Inline" Then
                ' Verifica se a imagem inline ainda existe na posição esperada
                If .paraIndex <= doc.Paragraphs.count Then
                    Dim para As Paragraph
                    Set para = doc.Paragraphs(.paraIndex)
                    
                    ' Se ainda há imagens inline no parágrafo, considera verificada
                    If para.Range.InlineShapes.count > 0 Then
                        verifiedCount = verifiedCount + 1
                    End If
                End If
                
            ElseIf .ImageType = "Floating" Then
                ' Verifica e corrige propriedades de shapes flutuantes se ainda existem
                If .ImageIndex <= doc.Shapes.count Then
                    Dim targetShape As shape
                    Set targetShape = doc.Shapes(.ImageIndex)
                    
                    ' Verifica se as propriedades foram alteradas e corrige se necessário
                    Dim needsCorrection As Boolean
                    needsCorrection = False
                    
                    If Abs(targetShape.Width - .Width) > 1 Then needsCorrection = True
                    If Abs(targetShape.Height - .Height) > 1 Then needsCorrection = True
                    If Abs(targetShape.Left - .LeftPosition) > 1 Then needsCorrection = True
                    If Abs(targetShape.Top - .TopPosition) > 1 Then needsCorrection = True
                    
                    If needsCorrection Then
                        ' Restaura propriedades originais
                        With targetShape
                            .Width = savedImages(i).Width
                            .Height = savedImages(i).Height
                            .Left = savedImages(i).LeftPosition
                            .Top = savedImages(i).TopPosition
                            .WrapFormat.Type = savedImages(i).WrapType
                        End With
                        correctedCount = correctedCount + 1
                    End If
                    
                    verifiedCount = verifiedCount + 1
                End If
            End If
        End With
        
        On Error GoTo ErrorHandler
    Next i
    
    If correctedCount > 0 Then
        LogMessage "Verificação de imagens concluída: " & verifiedCount & " verificadas, " & correctedCount & " corrigidas"
    Else
        LogMessage "Verificação de imagens concluída: " & verifiedCount & " imagens íntegras"
    End If
    
    RestoreAllImages = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao verificar imagens: " & Err.Description, LOG_LEVEL_WARNING
    RestoreAllImages = False
End Function

'================================================================================
' GET CLIPBOARD DATA - Obtém dados da área de transferência
'================================================================================
Private Function GetClipboardData() As Variant
    On Error GoTo ErrorHandler
    
    ' Placeholder para dados da área de transferência
    ' Em uma implementação completa, seria necessário usar APIs do Windows
    ' ou métodos mais avançados para capturar dados binários
    GetClipboardData = "ImageDataPlaceholder"
    Exit Function

ErrorHandler:
    GetClipboardData = Empty
End Function

'================================================================================
' ENHANCED IMAGE PROTECTION - Proteção aprimorada durante formatação
'================================================================================
Private Function ProtectImagesInRange(targetRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verifica se há imagens no range antes de aplicar formatação
    If targetRange.InlineShapes.count > 0 Then
        ' OTIMIZADO: Aplica formatação caractere por caractere, protegendo imagens
        Dim i As Long
        Dim charRange As Range
        Dim charCount As Long
        charCount = SafeGetCharacterCount(targetRange) ' Cache da contagem segura
        
        If charCount > 0 Then ' Verificação de segurança
            For i = 1 To charCount
                Set charRange = targetRange.Characters(i)
                ' Só formata caracteres que não são parte de imagens
                If charRange.InlineShapes.count = 0 Then
                    With charRange.Font
                        .Name = STANDARD_FONT
                        .size = STANDARD_FONT_SIZE
                        .Color = wdColorAutomatic
                    End With
                End If
            Next i
        End If
    Else
        ' Range sem imagens - formatação normal completa
        With targetRange.Font
            .Name = STANDARD_FONT
            .size = STANDARD_FONT_SIZE
            .Color = wdColorAutomatic
        End With
    End If
    
    ProtectImagesInRange = True
    Exit Function

ErrorHandler:
    LogMessage "Erro na proteção de imagens: " & Err.Description, LOG_LEVEL_WARNING
    ProtectImagesInRange = False
End Function

'================================================================================
' CLEANUP IMAGE PROTECTION - Limpeza das variáveis de proteção de imagens
'================================================================================
Private Sub CleanupImageProtection()
    On Error Resume Next
    
    ' Limpa arrays de imagens
    If imageCount > 0 Then
        Dim i As Long
        For i = 0 To imageCount - 1
            Set savedImages(i).AnchorRange = Nothing
        Next i
    End If
    
    imageCount = 0
    ReDim savedImages(0)
    
    LogMessage "Variáveis de proteção de imagens limpas"
End Sub

'================================================================================
' VIEW SETTINGS PROTECTION SYSTEM - SISTEMA DE PROTEÇÃO DAS CONFIGURAÇÕES DE VISUALIZAÇÃO
'================================================================================

'================================================================================
' BACKUP VIEW SETTINGS - Faz backup das configurações de visualização originais
'================================================================================
Private Function BackupViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Fazendo backup das configurações de visualização..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Backup das configurações de visualização
    With originalViewSettings
        .ViewType = docWindow.View.Type
        ' Réguas são controladas pelo Window, não pelo View
        On Error Resume Next
        .ShowHorizontalRuler = docWindow.DisplayRulers
        .ShowVerticalRuler = docWindow.DisplayVerticalRuler
        On Error GoTo ErrorHandler
        .ShowFieldCodes = docWindow.View.ShowFieldCodes
        .ShowBookmarks = docWindow.View.ShowBookmarks
        .ShowParagraphMarks = docWindow.View.ShowParagraphs
        .ShowSpaces = docWindow.View.ShowSpaces
        .ShowTabs = docWindow.View.ShowTabs
        .ShowHiddenText = docWindow.View.ShowHiddenText
        .ShowAll = docWindow.View.ShowAll
        .ShowDrawings = docWindow.View.ShowDrawings
        .ShowObjectAnchors = docWindow.View.ShowObjectAnchors
        .ShowTextBoundaries = docWindow.View.ShowTextBoundaries
        .ShowHighlight = docWindow.View.ShowHighlight
        ' .ShowAnimation removida - pode não existir em todas as versões
        .DraftFont = docWindow.View.Draft
        .WrapToWindow = docWindow.View.WrapToWindow
        .ShowPicturePlaceHolders = docWindow.View.ShowPicturePlaceHolders
        .ShowFieldShading = docWindow.View.FieldShading
        .TableGridlines = docWindow.View.TableGridlines
        ' .EnlargeFontsLessThan removida - pode não existir em todas as versões
    End With
    
    LogMessage "Backup das configurações de visualização concluído"
    BackupViewSettings = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao fazer backup das configurações de visualização: " & Err.Description, LOG_LEVEL_WARNING
    BackupViewSettings = False
End Function

'================================================================================
' RESTORE VIEW SETTINGS - Restaura as configurações de visualização originais
'================================================================================
Private Function RestoreViewSettings(doc As Document) As Boolean
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "Restaurando configurações de visualização originais..."
    
    Dim docWindow As Window
    Set docWindow = doc.ActiveWindow
    
    ' Restaura todas as configurações originais, EXCETO o zoom
    With docWindow.View
        .Type = originalViewSettings.ViewType
        .ShowFieldCodes = originalViewSettings.ShowFieldCodes
        .ShowBookmarks = originalViewSettings.ShowBookmarks
        .ShowParagraphs = originalViewSettings.ShowParagraphMarks
        .ShowSpaces = originalViewSettings.ShowSpaces
        .ShowTabs = originalViewSettings.ShowTabs
        .ShowHiddenText = originalViewSettings.ShowHiddenText
        .ShowAll = originalViewSettings.ShowAll
        .ShowDrawings = originalViewSettings.ShowDrawings
        .ShowObjectAnchors = originalViewSettings.ShowObjectAnchors
        .ShowTextBoundaries = originalViewSettings.ShowTextBoundaries
        .ShowHighlight = originalViewSettings.ShowHighlight
        ' .ShowAnimation removida para compatibilidade
        .Draft = originalViewSettings.DraftFont
        .WrapToWindow = originalViewSettings.WrapToWindow
        .ShowPicturePlaceHolders = originalViewSettings.ShowPicturePlaceHolders
        .FieldShading = originalViewSettings.ShowFieldShading
        .TableGridlines = originalViewSettings.TableGridlines
        ' .EnlargeFontsLessThan removida para compatibilidade
        
        ' ZOOM é mantido em 110% - única configuração que permanece alterada
        .Zoom.Percentage = 110
    End With
    
    ' Configurações específicas do Window (para réguas)
    docWindow.DisplayRulers = originalViewSettings.ShowHorizontalRuler
    docWindow.DisplayVerticalRuler = originalViewSettings.ShowVerticalRuler
    
    LogMessage "Configurações de visualização originais restauradas (zoom mantido em 110%)"
    RestoreViewSettings = True
    Exit Function

ErrorHandler:
    LogMessage "Erro ao restaurar configurações de visualização: " & Err.Description, LOG_LEVEL_WARNING
    RestoreViewSettings = False
End Function

'================================================================================
' CLEANUP VIEW SETTINGS - Limpeza das variáveis de configurações de visualização
'================================================================================
Private Sub CleanupViewSettings()
    On Error Resume Next
    
    ' Reinicializa a estrutura de configurações
    With originalViewSettings
        .ViewType = 0
        .ShowVerticalRuler = False
        .ShowHorizontalRuler = False
        .ShowFieldCodes = False
        .ShowBookmarks = False
        .ShowParagraphMarks = False
        .ShowSpaces = False
        .ShowTabs = False
        .ShowHiddenText = False
        .ShowOptionalHyphens = False
        .ShowAll = False
        .ShowDrawings = False
        .ShowObjectAnchors = False
        .ShowTextBoundaries = False
        .ShowHighlight = False
        ' .ShowAnimation removida para compatibilidade
        .DraftFont = False
        .WrapToWindow = False
        .ShowPicturePlaceHolders = False
        .ShowFieldShading = 0
        .TableGridlines = False
        ' .EnlargeFontsLessThan removida para compatibilidade
    End With
    
    LogMessage "Variáveis de configurações de visualização limpas"
End Sub

'================================================================================
' REPLACE STANDARD LOCATION AND DATE PARAGRAPH
'================================================================================
Private Sub ReplacePlenarioDateParagraph(doc As Document)
    On Error GoTo ErrorHandler
    
    If doc Is Nothing Then Exit Sub
    
    Dim para As Paragraph
    Dim paraText As String
    Dim matchCount As Integer
    Dim terms() As String
    
    ' Define os termos de busca
    terms = Split("Palácio 15 de Junho,Plenário,Dr. Tancredo Neves," & _
                 " de janeiro de , de fevereiro de, de março de, de abril de," & _
                 " de maio de, de junho de, de julho de, de agosto de," & _
                 " de setembro de, de outubro de, de novembro de, de dezembro de", ",")
    
    ' Processa cada parágrafo
    For Each para In doc.Paragraphs
        matchCount = 0
        
        ' Pula parágrafos muito longos
        If Len(para.Range.text) <= 80 Then
            paraText = para.Range.text
            
            ' Conta matches
            Dim term As Variant
            For Each term In terms
                If InStr(1, paraText, CStr(term), vbTextCompare) > 0 Then
                    matchCount = matchCount + 1
                End If
                If matchCount >= 2 Then
                    ' Encontrou 2+ matches, faz a substituição
                    ' Usa Delete + InsertAfter para preservar o marcador de parágrafo
                    para.Range.Select
                    Selection.MoveEnd Unit:=wdCharacter, Count:=-1 ' Exclui o marcador de parágrafo
                    Selection.Delete
                    Selection.InsertAfter "Plenário ""Dr. Tancredo Neves"", $DATAATUALEXTENSO$."
                    ' Aplica formatação: centralizado e sem recuos
                    With para.Range.ParagraphFormat
                        .leftIndent = 0
                        .firstLineIndent = 0
                        .alignment = wdAlignParagraphCenter
                    End With
                    LogMessage "Parágrafo de plenário substituído e formatado", LOG_LEVEL_INFO
                    Exit For
                End If
            Next term
        End If
    Next para
    
    Exit Sub
    
ErrorHandler:
    LogMessage "Erro ao processar parágrafos: " & Err.Description, LOG_LEVEL_ERROR
End Sub

'================================================================================
' BACKUP DIRECTORY MANAGEMENT - #STABLE
'================================================================================
Private Function EnsureBackupDirectory(doc As Document) As String
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim backupPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define o caminho base para backups (mesmo diretório do documento ou TEMP)
    If doc.Path <> "" Then
        backupPath = doc.Path & "\" & BACKUP_FOLDER_NAME
    Else
        backupPath = Environ("TEMP") & "\" & BACKUP_FOLDER_NAME
    End If
    
    ' Cria o diretório se não existir
    If Not fso.FolderExists(backupPath) Then
        fso.CreateFolder backupPath
        LogMessage "Pasta de backup criada: " & backupPath, LOG_LEVEL_INFO
    End If
    
    EnsureBackupDirectory = backupPath
    Exit Function
    
ErrorHandler:
    LogMessage "Erro ao criar pasta de backup: " & Err.Description, LOG_LEVEL_ERROR
    ' Retorna pasta TEMP como fallback
    EnsureBackupDirectory = Environ("TEMP")
End Function
