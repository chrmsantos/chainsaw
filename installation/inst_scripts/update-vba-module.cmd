@echo off
REM =============================================================================
REM CHAINSAW - Atualizador de Modulo VBA (Wrapper)
REM =============================================================================
REM Este arquivo agora redireciona para o launcher unificado chainsaw.cmd
REM Mantido para compatibilidade com scripts existentes
REM =============================================================================

"%~dp0chainsaw.cmd" update-vba %*
exit /b %ERRORLEVEL%
