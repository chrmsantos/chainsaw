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

' Semantic version of the Chainsaw modular toolkit
Public Const CHAINSAW_VERSION As String = "1.0.0-modular"

' Public entry point replacing direct calls to StandardizeDocumentMain.
' Kept thin for stability; delegates to existing implementation in the
' original monolithic module until functions are migrated.
Public Sub ChainsawRun()
    ' New modular pipeline orchestration (deprecated monolith call removed)
    On Error GoTo OrchestratorFatal
    RunChainsawPipeline False
    ' Optional: auto smoke test if enabled in configuration (after first run)
    If Config.debugMode And Config.autoRunSmokeTest Then
        On Error Resume Next
        ChainsawSmokeTest
        On Error GoTo OrchestratorFatal
    End If
    Exit Sub
OrchestratorFatal:
    On Error Resume Next
    Application.StatusBar = "Chainsaw fatal error: " & Err.Number & " - " & Err.Description
End Sub

' Transitional alias so users can start using a clearer macro name.
Public Sub Chainsaw()
    ChainsawRun
End Sub

' Placeholder stubs for future modular layers (to be filled in as code is migrated):
Public Sub ChainsawLoadConfiguration()
    RunChainsawPipeline True, True, False, False
End Sub

Public Sub ChainsawValidateActiveDocument()
    RunChainsawPipeline True, True, True, False
End Sub

Public Sub ChainsawFormatActiveDocument()
    RunChainsawPipeline True, True, True, True
End Sub

' =================================================================================
' INTERNAL PIPELINE (modular)
' =================================================================================
Private Sub RunChainsawPipeline(Optional loadConfig As Boolean = True, _
                                Optional performValidation As Boolean = True, _
                                Optional includeHeavyValidation As Boolean = True, _
                                Optional applyFormatting As Boolean = True)
    On Error GoTo PipelineFail
    Dim doc As Document
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo PipelineFail
    If doc Is Nothing Then
        Application.StatusBar = "No active document"
        Exit Sub
    End If

    ' 1. Configuration
    If loadConfig Then
        If Not LoadConfiguration() Then
            Application.StatusBar = "Config load failed (defaults used)"
        End If
    End If

    ' 2. Basic validations (subset of monolith validations)
    If performValidation Then
        If Config.CheckWordVersion Then
            If Not CheckWordVersion() Then
                Application.StatusBar = "Word version unsupported"
                Exit Sub
            End If
        End If
        If Not EnsureDocumentEditable(doc) Then
            Application.StatusBar = "Document not editable"
            Exit Sub
        End If
        If Config.ValidateDocumentIntegrity Then
            If Not ValidateDocumentIntegrity(doc) Then
                Application.StatusBar = "Integrity validation failed"
                Exit Sub
            End If
        End If
        If includeHeavyValidation Then
            If Config.ValidatePropositionType Then Call ValidatePropositionType(doc)
            If Config.ValidateContentConsistency Then Call ValidateContentConsistency(doc)
        End If
    End If

    ' 3. Performance optimization
    Call InitializePerformanceOptimization

    ' 4. Formatting (delegates to modFormatting functions based on config flags)
    If applyFormatting Then
    If Config.ApplyPageSetup Then Call modFormatting.ApplyPageSetup(doc)
    If Config.applyStandardFont Then Call modFormatting.ApplyStdFont(doc)
    If Config.applyStandardParagraphs Then Call modFormatting.ApplyStdParagraphs(doc)
    If Config.FormatFirstParagraph Then Call modFormatting.FormatFirstParagraph(doc)
    If Config.FormatSecondParagraph Then Call modFormatting.FormatSecondParagraph(doc)
    If Config.FormatConsiderandoParagraphs Then Call modFormatting.FormatConsiderandoParagraphs(doc)
    If Config.formatJustificativaParagraphs Then Call modFormatting.FormatJustificativaAnexoParagraphs(doc)
    If Config.FormatNumberedParagraphs Then Call modFormatting.FormatNumberedParagraphs(doc)
    If Config.EnableHyphenation Then Call modFormatting.EnableHyphenation(doc)
    If Config.RemoveWatermark Then Call modFormatting.RemoveWatermark(doc)
    If Config.InsertHeaderstamp Then Call modFormatting.InsertHeaderstamp(doc)
    If Config.InsertFooterstamp Then Call modFormatting.InsertFooterstamp(doc)
    If Config.ApplyTextReplacements Then Call modFormatting.ApplyTextReplacements(doc)
    If Config.ApplySpecificParagraphReplacements Then Call modFormatting.ApplySpecificParagraphReplacements(doc)
        ' Light cleanups if present in monolith utilities (optional):
        If Config.CleanMultipleSpaces Then Call CleanMultipleSpaces(doc)
    End If

    Application.StatusBar = "Chainsaw pipeline completed (v" & CHAINSAW_VERSION & ")"
    Exit Sub

PipelineFail:
    On Error Resume Next
    Application.StatusBar = "Pipeline error: " & Err.Number
End Sub
