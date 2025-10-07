Attribute VB_Name = "modLog"
'================================================================================
' MODULE: modLog
' PURPOSE: Minimal no-op logging stubs to preserve call structure without overhead.
'================================================================================
Option Explicit

Public Sub LogStepStart(message As String)
    ' Intentionally no-op (placeholder for future instrumentation)
End Sub

Public Sub LogStepEnd(success As Boolean)
    ' Intentionally no-op
End Sub

Public Sub LogInfo(message As String)
    ' Intentionally no-op
End Sub

Public Sub LogError(message As String)
    ' Intentionally no-op
End Sub
