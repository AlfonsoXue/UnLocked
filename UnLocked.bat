@echo off
rem ==========================================================================
rem  UnLocked launcher
rem  - Double-click to open the GUI.
rem  - Drag files or folders onto this .bat to add them to the list.
rem ==========================================================================
setlocal
set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PSEXE%" set "PSEXE=powershell.exe"
"%PSEXE%" -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0src\UnLocked.ps1" %*
if errorlevel 2 pause
endlocal
