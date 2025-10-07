' SPDX-License-Identifier: GPL-3.0-or-later
' =============================================================================
' MODULE: ChainsawOrchestrator
' PURPOSE: High-level orchestration entry points decoupled from monolithic logic.
'          This is Stage 1 of modularization. Subsequent stages will migrate
'          configuration, validation, formatting, replacements, and cleanup
'          routines into dedicated modules while keeping this orchestrator stable.
' =============================================================================
Option Explicit
Option Private Module

' Public entry point replacing direct calls to StandardizeDocumentMain.
' Kept thin for stability; delegates to existing implementation in the
' original monolithic module until functions are migrated.
Public Sub ChainsawRun()
    ' Backward compatibility: reuse existing standardized pipeline.
    ' Future: call into modular layers (Config.Load, Validation.RunAll, Formatting.Apply, etc.)
    On Error GoTo OrchestratorFatal
    StandardizeDocumentMain
    Exit Sub
OrchestratorFatal:
    ' Fail-soft: basic status bar notice; detailed recovery handled by underlying code.
    On Error Resume Next
    Application.StatusBar = "Chainsaw Orchestrator fatal error: " & Err.Number & " - " & Err.Description
End Sub

' Transitional alias so users can start using a clearer macro name.
Public Sub Chainsaw()
    ChainsawRun
End Sub

' Placeholder stubs for future modular layers (to be filled in as code is migrated):
Public Sub ChainsawLoadConfiguration()
    ' Will delegate to Config module once extracted.
    StandardizeDocumentMain ' temporary minimal reuse; will be refactored
End Sub

Public Sub ChainsawFormatActiveDocument()
    ' Will eventually just run formatting phase.
    StandardizeDocumentMain
End Sub

Public Sub ChainsawValidateActiveDocument()
    ' Will eventually run only validation subset.
    StandardizeDocumentMain
End Sub
