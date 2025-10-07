' SPDX-License-Identifier: GPL-3.0-or-later
' =============================================================================
' MODULE: modUtils
' PURPOSE: Shared utility functions (string normalization, placeholder replacement,
'          timing helpers, unit conversions) extracted from monolithic module.
' =============================================================================
Option Explicit

' Timing (shared start time reference lives in original orchestrator/monolith)
Public Function ElapsedSeconds(ByVal processingStartTime As Single) As Long
    On Error Resume Next
    If processingStartTime <= 0 Then
        ElapsedSeconds = 0
    Else
        ElapsedSeconds = CLng(Timer - processingStartTime)
        If ElapsedSeconds < 0 Then
            ElapsedSeconds = ElapsedSeconds + 86400& ' handle midnight rollover
        End If
    End If
End Function

' Unit conversion: points to centimeters (inverse of Word native CentimetersToPoints)
Public Function CmFromPoints(ByVal pts As Double) As Double
    CmFromPoints = (pts * 2.54#) / 72#
End Function

' UI STRING NORMALIZATION - produce ASCII-safe text for MsgBox dialogs
Public Function NormalizeForUI(ByVal s As String) As String
    On Error Resume Next
    If Not dialogAsciiNormalizationEnabled Then
        NormalizeForUI = s
        Exit Function
    End If
    Dim i As Long, ch As String, code As Long, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        Select Case code
            Case 192 To 197, 224 To 229: out = out & "a"
            Case 199: out = out & "C"
            Case 231: out = out & "c"
            Case 200 To 203, 232 To 235: out = out & "e"
            Case 204 To 207, 236 To 239: out = out & "i"
            Case 210 To 214, 242 To 246: out = out & "o"
            Case 217 To 220, 249 To 252: out = out & "u"
            Case 209: out = out & "N"
            Case 241: out = out & "n"
            Case 8211, 8212: out = out & "-"
            Case 8216, 8217: out = out & "'"
            Case 8220, 8221, 171, 187: out = out & Chr$(34)
            Case 10, 13: out = out & ch
            Case 32 To 126: out = out & ch
            Case Else: out = out & "?"
        End Select
    Next i
    NormalizeForUI = out
End Function

' LOG STRING NORMALIZATION (retained even though logging removed - may be useful later)
Public Function NormalizeForLog(ByVal s As String) As String
    On Error Resume Next
    Dim i As Long, ch As String, code As Long, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code >= 32 And code <= 126 Then
            out = out & ch
        ElseIf ch = vbCr Or ch = vbLf Or ch = vbTab Then
            out = out & ch
        Else
            out = out & "?"
        End If
    Next i
    NormalizeForLog = out
End Function

' ReplacePlaceholders - convenience wrapper for {{KEY}} replacements
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
