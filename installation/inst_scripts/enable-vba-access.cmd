@echo off
REM =============================================================================
REM CHAINSAW - Habilitador de Acesso VBA (Wrapper)
REM =============================================================================
REM Este arquivo agora redireciona para o launcher unificado chainsaw.cmd
REM Mantido para compatibilidade com scripts existentes
REM =============================================================================

"%~dp0chainsaw.cmd" enable-vba %*
exit /b %ERRORLEVEL%
