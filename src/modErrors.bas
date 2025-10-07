'Attribute VB_Name = "modErrors"
'================================================================================
' MODULE: modErrors
' PURPOSE: Centralized lightweight error reporting & status updates.
' NOTE: Designed to be minimal (no file I/O) per beta objectives.
'================================================================================
Option Explicit

' Report a handled, expected validation or user-facing warning.
Public Sub ReportWarning(context As String, message As String)
    On Error Resume Next
    If Len(message) = 0 Then Exit Sub
    If Config.showStatusBarUpdates Then Application.StatusBar = "Chainsaw: " & Left(NormalizeForUI(message), 180)
End Sub

' Report an unexpected runtime error. Fails open (no hard stop) unless critical.
Public Sub ReportUnexpected(procName As String, Optional extra As String = "")
    On Error Resume Next
    Dim msg As String
    msg = "Erro inesperado em " & procName & IIf(extra <> "", ": " & extra, "")
    If Err.Number <> 0 Then
        msg = msg & " (" & Err.Number & ": " & Err.Description & ")"
    End If
    If Config.showStatusBarUpdates Then Application.StatusBar = "Chainsaw: " & Left(NormalizeForUI(msg), 180)
    ' Optional user dialog (honors debugMode)
    If Config.debugMode And Config.showProgressMessages Then
        MsgBox NormalizeForUI(msg), vbExclamation, NormalizeForUI("Chainsaw - Erro")
    End If
End Sub

' Clear status bar when finishing if enabled.
Public Sub ReportCompletion(success As Boolean)
    On Error Resume Next
    If Not Config.showStatusBarUpdates Then Exit Sub
    If success Then
        Application.StatusBar = "Chainsaw: concluído"
    Else
        Application.StatusBar = "Chainsaw: concluído com avisos"
    End If
End Sub
