Attribute VB_Name = "modUI"
Option Explicit
'
' MODULE: modUI
' PURPOSE: Centralized lightweight UI string helpers shared across modules.
' NOTES:
'  - Central utility functions NormalizeForUI / ReplacePlaceholders.
'  - Public exposure ensures validation, pipeline, and error messaging modules can
'    use consistent transformation logic.
'  - ASCII normalization is controlled by a module constant; adjust if future
'    configuration toggling is desired.
'
Private Const DIALOG_ASCII_NORMALIZATION_ENABLED As Boolean = True

' NormalizeForUI - produce ASCII-safe text for MsgBox dialogs.
' Replaces common accented / curly punctuation characters with simpler ASCII so
' dialogs render consistently across varied Windows regional settings.
Public Function NormalizeForUI(ByVal s As String) As String
    On Error Resume Next
    If Not DIALOG_ASCII_NORMALIZATION_ENABLED Then
        NormalizeForUI = s
        Exit Function
    End If
    Dim i As Long, ch As String, code As Long
    Dim out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 192 To 197, 224 To 229: out = out & "a"   ' ÀÁÂÃÄÅ àáâãäå
            Case 199: out = out & "C"                      ' Ç
            Case 231: out = out & "c"                      ' ç
            Case 200 To 203, 232 To 235: out = out & "e"
            Case 204 To 207, 236 To 239: out = out & "i"
            Case 210 To 214, 242 To 246: out = out & "o"
            Case 217 To 220, 249 To 252: out = out & "u"
            Case 209: out = out & "N"                      ' Ñ
            Case 241: out = out & "n"                      ' ñ
            Case 8211, 8212: out = out & "-"               ' en/em dash
            Case 8216, 8217: out = out & "'"               ' curly apostrophes
            Case 8220, 8221, 171, 187: out = out & """"    ' various quotes -> standard quote
            Case 10, 13: out = out & ch                    ' CR/LF
            Case 32 To 126: out = out & ch                 ' standard ASCII
            Case Else: out = out & "?"
        End Select
    Next i
    NormalizeForUI = out
End Function

' ReplacePlaceholders - convenience wrapper for common {{KEY}} replacements.
' Example: ReplacePlaceholders(MSG_ERR_VERSION, "MIN", 14, "CUR", Application.Version)
Public Function ReplacePlaceholders(ByVal template As String, ParamArray kv()) As String
    On Error Resume Next
    Dim i As Long, result As String, k As String, val As String
    result = template
    For i = LBound(kv) To UBound(kv) Step 2
        If i + 1 <= UBound(kv) Then
            k = CStr(kv(i))
            val = CStr(kv(i + 1))
            result = Replace(result, "{{" & k & "}}", val)
        End If
    Next i
    ReplacePlaceholders = result
End Function
