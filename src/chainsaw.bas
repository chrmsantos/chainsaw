Attribute VB_Name = "chainsaw"
' =============================================================================
' PROJECT: CHAINSAW PROPOSITURAS
' FILE: chainsaw.bas (public entry stub ONLY)
' =============================================================================
' Purpose: Maintains historical macro name and delegates to the pipeline.
'          ALL logic lives in specialized modules. Keep this file minimal.
' =============================================================================
' License: Modified Apache 2.0 (see LICENSE)
' Version: 1.0.0-Beta2
' =============================================================================
Option Explicit

' Public entry point retained for backward compatibility (old macro name)
Public Sub ChainsawProcess()
    Dim ok As Boolean
    ok = RunChainsawPipeline()
    If ok Then
        Application.StatusBar = "Chainsaw: processamento concluido"
    Else
        Application.StatusBar = "Chainsaw: processamento falhou"
    End If
End Sub

' =============================================================================
' NOTE:
' This file was intentionally truncated on 2025-10-06. All previous private
' helper routines (formatting, validation, cleanup, numbering, backups, view
' configuration, visual element removal, save/exit helpers, etc.) were migrated
' or superseded by implementations in the following modules:
'   modPipeline
'   modFormatting
'   modReplacements
'   modValidation
'   modSafety
'   modConfig
'   modMessages
'   modConstants
'   modErrors
'   modSelfTest
'   modUI
' Historic full content preserved at: legacy_chainsaw_snapshot.bas
' Do NOT reintroduce logic here—keep stub minimal to avoid divergence.
' =============================================================================

' End of file – keep clean.
