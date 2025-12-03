@echo off
chcp 65001 >nul
:: ============================================================================
:: CHAINSAW - Atualizacao Automatica via GitHub (Wrapper CMD)
:: ============================================================================
:: Versao: 1.0.0
:: Licenca: GNU GPLv3
:: ============================================================================

echo.
echo ═══════════════════════════════════════════════════════════════
echo    CHAINSAW - ATUALIZACAO AUTOMATICA VIA GITHUB
echo ═══════════════════════════════════════════════════════════════
echo.

:: Verificar se PowerShell esta disponivel
where powershell.exe >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] PowerShell nao encontrado!
    echo.
    pause
    exit /b 1
)

:: Executar script PowerShell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0update-from-github.ps1" %*

:: Verificar resultado
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [OK] Atualizacao concluida com sucesso!
) else (
    echo.
    echo [ERRO] Atualizacao falhou - verifique as mensagens acima
)

echo.
pause
