@echo off
REM =============================================================================
REM CHAINSAW - Exportador de Configuracoes (Wrapper)
REM =============================================================================
REM Este arquivo agora redireciona para o launcher unificado chainsaw.cmd
REM Mantido para compatibilidade com scripts existentes
REM =============================================================================

"%~dp0chainsaw.cmd" export %*
exit /b %ERRORLEVEL%
