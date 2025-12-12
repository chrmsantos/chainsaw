@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%importar_configs.ps1"

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
if %EXIT_CODE% equ 0 (
    for /f "usebackq delims=" %%L in (`powershell -NoProfile -NoLogo -Command "try { $root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path); $logDir = Join-Path $root 'exported-config\\logs'; if (Test-Path $logDir) { $l = Get-ChildItem -Path $logDir -Filter 'import_*.log' -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($l) { $l.FullName } } } catch { }"`) do (
        echo Log: %%L
    )
)
goto :maybe_pause

echo EXITCODE=%EXIT_CODE%

:pause_fail
set "EXIT_CODE=1"

:maybe_pause
if not defined CHAINSAW_NO_PAUSE (
    echo %CMDCMDLINE% ^| find /I "%~0" >nul
    if not errorlevel 1 (
        echo.
        pause
    )
)

endlocal
exit /b %EXIT_CODE%
