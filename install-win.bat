@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "SCOPE=User"
set "SKIP_DEBUG=0"

:parse_args
if "%~1"=="" goto parsed
if /I "%~1"=="-h" goto usage
if /I "%~1"=="--help" goto usage

if /I "%~1"=="--scope" (
    if "%~2"=="" (
        echo Missing value for --scope. Use User or System.
        exit /b 1
    )
    set "SCOPE=%~2"
    shift
    shift
    goto parse_args
)

if /I "%~1"=="--skip-debug" (
    set "SKIP_DEBUG=1"
    shift
    goto parse_args
)

echo Unknown argument: %~1
exit /b 1

:parsed
if /I not "%SCOPE%"=="User" if /I not "%SCOPE%"=="System" (
    echo Invalid --scope value: %SCOPE%
    echo Allowed values: User or System
    exit /b 1
)

if /I "%SCOPE%"=="System" (
    net session >nul 2>&1
    if errorlevel 1 (
        echo System scope requires an elevated Command Prompt.
        exit /b 1
    )
)

for %%I in ("%~dp0.") do set "REPO_ROOT=%%~fI"
set "EXT_NAME=PremiereGridMaker"

if /I "%SCOPE%"=="System" (
    set "BASE_PATH=%ProgramFiles(x86)%\Common Files\Adobe\CEP\extensions"
) else (
    set "BASE_PATH=%APPDATA%\Adobe\CEP\extensions"
)

set "INSTALL_PATH=%BASE_PATH%\%EXT_NAME%"

if not exist "%BASE_PATH%" mkdir "%BASE_PATH%"
if exist "%INSTALL_PATH%" rmdir /s /q "%INSTALL_PATH%"
mkdir "%INSTALL_PATH%"

robocopy "%REPO_ROOT%" "%INSTALL_PATH%" /E /R:2 /W:1 /NFL /NDL /NJH /NJS /NP /XD ".git" "node_modules"
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
    echo Installation failed during file copy. Robocopy exit code: %RC%
    exit /b %RC%
)

if "%SKIP_DEBUG%"=="0" (
    for %%V in (8 9 10 11) do (
        reg add "HKCU\Software\Adobe\CSXS.%%V" /v PlayerDebugMode /t REG_SZ /d 1 /f >nul
    )
)

echo Installed "%EXT_NAME%" to: %INSTALL_PATH%
if "%SKIP_DEBUG%"=="1" (
    echo Skipped CEP debug mode changes.
) else (
    echo CEP debug mode enabled for CSXS.8 to CSXS.11 ^(HKCU^).
)
echo Open Premiere Pro: Window ^> Extensions ^(Legacy^) ^> Premiere Grid Maker
exit /b 0

:usage
echo Usage: install-win.bat [--scope User^|System] [--skip-debug]
exit /b 0
