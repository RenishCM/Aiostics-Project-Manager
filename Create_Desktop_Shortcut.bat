@echo off
setlocal enabledelayedexpansion

:: ─────────────────────────────────────────────────────────────
::  Aiostic Project Manager — Desktop Shortcut Creator
::  Copyright (c) 2025 Renish C M
:: ─────────────────────────────────────────────────────────────

:: Get the folder this .bat lives in (strip trailing backslash)
set "DIR=%~dp0"
set "DIR=%DIR:~0,-1%"

set "HTML=%DIR%\Aiostics_Project_Manager.html"
set "ICO=%DIR%\aiostics.ico"

:: Check files exist before proceeding
if not exist "%HTML%" (
  echo.
  echo  ERROR: Aiostics_Project_Manager.html not found in this folder.
  echo  Make sure all files are in the same folder before running this.
  echo.
  pause
  exit /b 1
)

if not exist "%ICO%" (
  echo.
  echo  WARNING: aiostics.ico not found. Shortcut will use default browser icon.
  echo.
  set "ICO="
)

:: Create the shortcut
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$desktop = [System.Environment]::GetFolderPath('Desktop');" ^
  "$s = $ws.CreateShortcut($desktop + '\Aiostic Project Manager.lnk');" ^
  "$s.TargetPath = '%HTML%';" ^
  "$s.IconLocation = '%ICO%,0';" ^
  "$s.Description = 'Aiostic Project Manager — Offline Project Planner';" ^
  "$s.WindowStyle = 1;" ^
  "$s.Save();" ^
  "Write-Host 'Shortcut created successfully.'"

echo.
echo  Done! Aiostic Project Manager shortcut has been added to your Desktop.
echo  Right-click the Desktop and choose Refresh if you don't see it yet.
echo.
pause
