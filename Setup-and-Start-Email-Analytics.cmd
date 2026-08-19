@echo off
setlocal EnableExtensions
title Email Analytics Dashboard Setup
cd /d "%~dp0"

echo.
echo ================================================
echo   Email Analytics Dashboard - One-Click Setup
echo ================================================
echo.

if not exist "requirements.txt" (
  echo ERROR: Run this file from inside the email-analytics-dashboard folder.
  goto :failed
)

echo [1/6] Selecting Classic Outlook...
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Preferences" /v UseNewOutlook /t REG_DWORD /d 0 /f >nul 2>&1
if errorlevel 1 (
  echo ERROR: Windows could not select Classic Outlook for this user.
  goto :failed
)
reg add "HKCU\Software\Microsoft\Office\16.0\Outlook\Options\General" /v HideNewOutlookToggle /t REG_DWORD /d 1 /f >nul 2>&1

echo [2/6] Finding and opening Classic Outlook...
set "OUTLOOK_EXE="
for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" /ve 2^>nul ^| find /i "REG_SZ"') do set "OUTLOOK_EXE=%%B"
if not defined OUTLOOK_EXE for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" /ve 2^>nul ^| find /i "REG_SZ"') do set "OUTLOOK_EXE=%%B"
if not defined OUTLOOK_EXE if exist "%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE" set "OUTLOOK_EXE=%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE"
if not defined OUTLOOK_EXE if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\OUTLOOK.EXE" set "OUTLOOK_EXE=%ProgramFiles(x86)%\Microsoft Office\root\Office16\OUTLOOK.EXE"

if not defined OUTLOOK_EXE (
  echo ERROR: Classic Outlook is not installed. Install Outlook Classic and run this file again.
  goto :failed
)
if not exist "%OUTLOOK_EXE%" (
  echo ERROR: Classic Outlook was registered at an invalid location:
  echo %OUTLOOK_EXE%
  goto :failed
)

start "" "%OUTLOOK_EXE%"
timeout /t 5 /nobreak >nul

echo [3/6] Checking Python...
if exist ".venv\Scripts\python.exe" goto :install_dependencies

set "PYTHON_BOOTSTRAP="
where py.exe >nul 2>&1
if not errorlevel 1 set "PYTHON_BOOTSTRAP=py"
if not defined PYTHON_BOOTSTRAP (
  where python.exe >nul 2>&1
  if not errorlevel 1 set "PYTHON_BOOTSTRAP=python"
)
if not defined PYTHON_BOOTSTRAP (
  echo ERROR: Python 3 is not installed or is not available in PATH.
  echo Install Python 3, select "Add Python to PATH", and run this file again.
  goto :failed
)

echo [4/6] Creating the Python environment...
%PYTHON_BOOTSTRAP% -m venv ".venv"
if errorlevel 1 goto :python_failed

:install_dependencies
echo [5/6] Installing required packages...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :python_failed
".venv\Scripts\python.exe" -m pip install -r "requirements.txt"
if errorlevel 1 goto :python_failed

echo [6/6] Registering automatic startup and opening the dashboard...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$project=(Resolve-Path '.').Path; $pythonw=Join-Path $project '.venv\Scripts\pythonw.exe'; $startup=[Environment]::GetFolderPath('Startup'); $shortcut=Join-Path $startup 'Email Analytics Dashboard.lnk'; $shell=New-Object -ComObject WScript.Shell; $link=$shell.CreateShortcut($shortcut); $link.TargetPath=$pythonw; $link.Arguments='-m email_analytics.startup'; $link.WorkingDirectory=$project; $link.WindowStyle=7; $link.Save()"
if errorlevel 1 (
  echo ERROR: The Windows startup shortcut could not be created.
  goto :failed
)

".venv\Scripts\python.exe" -m email_analytics.startup
if errorlevel 1 goto :python_failed

echo.
echo Setup completed successfully.
echo The dashboard is available at http://127.0.0.1:8501
echo It will start automatically after future Windows sign-ins.
echo.
pause
exit /b 0

:python_failed
echo.
echo ERROR: Python setup or application startup failed.
echo Check your internet connection and try again.

:failed
echo.
echo Setup did not complete.
pause
exit /b 1
