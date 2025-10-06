Option Explicit
'================================================================================
' MODULE: modPipeline (formerly modMain)
' PURPOSE: Orchestrate the high-level processing pipeline while all concrete
'          formatting, replacement, validation, safety and messaging logic
'          resides in dedicated modules. This file should remain thin.
' NOTE:    Original heavy legacy content from modMain.bas has been migrated.
'          Keep ONLY orchestration & sequencing here.
'================================================================================

Public Function RunChainsawPipeline() As Boolean
    RunChainsawPipeline = modMain.RunChainsawPipeline() ' Temporary delegation until full internal migration.
End Function

' TODO: Inline the body of RunChainsawPipeline from modMain and then remove modMain.

' End of file.