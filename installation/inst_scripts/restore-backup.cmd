@echo off
REM =============================================================================
REM CHAINSAW - Restaurador de Backup (Wrapper)
REM =============================================================================
REM Este arquivo agora redireciona para o launcher unificado chainsaw.cmd
REM Mantido para compatibilidade com scripts existentes
REM =============================================================================

"%~dp0chainsaw.cmd" restore %*
exit /b %ERRORLEVEL%
