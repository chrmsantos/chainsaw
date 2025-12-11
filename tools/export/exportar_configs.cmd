@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%exportar_configs.ps1"

where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo [ERRO] PowerShell nao encontrado.
    goto :pause_fail
)

if not exist "%PS_SCRIPT%" (
    echo [ERRO] Script PowerShell nao encontrado: %PS_SCRIPT%
    goto :pause_fail
)

powershell.exe -ExecutionPolicy Bypass -NoProfile -NoLogo -File "%PS_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"
goto :maybe_pause

:pause_fail
set "EXIT_CODE=1"

:maybe_pause
echo %CMDCMDLINE% ^| find /I "%~0" >nul
if not errorlevel 1 (
    echo.
    pause
)

endlocal
exit /b %EXIT_CODE%
