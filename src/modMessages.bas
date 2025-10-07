Attribute VB_Name = "modMessages"
Option Explicit
'================================================================================
' MODULE: modMessages
' PURPOSE: Centralize ALL user-facing texts (messages & titles) with lightweight
'          placeholder support handled by ReplacePlaceholders(). Strings are
'          ASCII-hardened to guarantee VBA module export/import reliability in
'          Word 2010+ environments with mixed locale encoding.
' PLACEHOLDER STYLE: {{KEY}}  (ReplacePlaceholders takes KEY/value pairs)
'================================================================================

' ---- Messages (MSG_*) ----
Public Const MSG_UNSAVED As String = _
	"O documento possui alteracoes nao salvas." & vbCrLf & _
	"Deseja salvar antes de continuar?"

Public Const MSG_LARGE_DOC As String = _
	"O documento e grande ({{SIZE}} caracteres)." & vbCrLf & _
	"Continuar o processamento?"

Public Const MSG_VALIDATION_ERROR As String = _
	"Erro de validacao: {{ERR}}"

Public Const MSG_DOC_TYPE_WARNING As String = _
	"O tipo de proposicao pode estar incorreto ou inconsistente." & vbCrLf & _
	"Deseja continuar mesmo assim?"

Public Const MSG_PROCESSING_CANCELLED As String = _
	"Processamento cancelado pelo usuario."

Public Const MSG_SAVE_ERROR As String = _
	"Falha ao salvar o documento. Verifique permissao ou caminho."

Public Const MSG_OPERATION_CANCELLED As String = _
	"Operacao cancelada. Nenhuma mudanca adicional foi aplicada."

Public Const MSG_CRITICAL_SAVE_EXIT As String = _
	"Erro critico ao salvar e encerrar: {{ERR}}" & vbCrLf & _
	"Salve manualmente se possivel antes de fechar o Word."

Public Const MSG_INCONSISTENCY_WARNING As String = _
	"Possivel inconsistencia entre cabecalho e ementa ({{COMMON}} palavras em comum)." & vbCrLf & _
	"Trecho analisado: {{Ementa}}" & vbCrLf & _
	"Continuar mesmo assim?"

' Structural validation messages
Public Const MSG_EMPTY_DOC As String = _
	"Documento vazio ou sem paragrafo com texto identificavel."
Public Const MSG_PARAGRAPH_EXCESS As String = _
	"Documento possui {{COUNT}} paragrafos (muito acima do esperado)." & vbCrLf & _
	"Deseja continuar mesmo assim?"
Public Const MSG_FIRST_PARA_SHORT As String = _
	"Primeiro paragrafo util e muito curto. Confirmar continuidade?"

' ---- Titles (TITLE_*) ----
Public Const TITLE_LARGE_DOC As String = "Documento grande"
Public Const TITLE_UNSAVED As String = "Documento nao salvo"
Public Const TITLE_VALIDATION_ERROR As String = "Erro de validacao"
Public Const TITLE_DOC_TYPE As String = "Tipo de proposicao"
Public Const TITLE_OPERATION_CANCELLED As String = "Operacao cancelada"
Public Const TITLE_SAVE_ERROR As String = "Erro ao salvar"
Public Const TITLE_FINAL_CONFIRM As String = "Confirmar operacao"
Public Const TITLE_CRITICAL_SAVE_EXIT As String = "Erro critico"
Public Const TITLE_CONSISTENCY As String = "Consistencia de conteudo"

' ---- System Name ----
Public Const SYSTEM_NAME As String = "Chainsaw Proposituras"

' End of file.
