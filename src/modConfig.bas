'================================================================================
' MODULE: modConfig
' PURPOSE: Configuration loading and parsing for Chainsaw Proposituras.
' NOTE: Extracted from monolithic chainsaw.bas on 2025-10-06.
'================================================================================
Option Explicit

Private Type ChainsawConfiguration
    headerImagePath As String
End Type

Public Config As ChainsawConfiguration

Public Function LoadConfiguration() As Boolean
    LoadConfiguration = True ' Placeholder: logic remains in original until fully migrated
End Function
