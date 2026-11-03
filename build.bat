@echo off
rem ==========================================================================
rem  Build dist\UnLocked.exe
rem  Uses the C# compiler that ships with Windows (.NET Framework 4.x), so no
rem  Visual Studio or .NET SDK is required.  In VS Code press Ctrl+Shift+B.
rem ==========================================================================
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo [ERROR] csc.exe not found. .NET Framework 4.x is required.
  exit /b 1
)

set "SMA="
for /f "delims=" %%F in ('dir /b /s "%WINDIR%\Microsoft.NET\assembly\GAC_MSIL\System.Management.Automation\System.Management.Automation.dll" 2^>nul') do set "SMA=%%F"
if not defined SMA (
  echo [ERROR] System.Management.Automation.dll not found. Windows PowerShell 5.1 is required.
  exit /b 1
)

rem Embed every script in src\ as a resource named UnLocked.src.<file name>
set "RES="
for %%F in (src\*.ps1 src\*.cs) do set "RES=!RES! /resource:%%F,UnLocked.src.%%~nxF"

if not exist dist mkdir dist
echo Building dist\UnLocked.exe ...
"%CSC%" /nologo /target:winexe /platform:anycpu /optimize+ /utf8output ^
  /out:dist\UnLocked.exe ^
  /win32icon:build\UnLocked.ico /win32manifest:build\UnLocked.manifest ^
  /reference:"%SMA%" /reference:System.Windows.Forms.dll ^
  !RES! ^
  build\UnLockedLauncher.cs build\AssemblyInfo.cs
if errorlevel 1 (
  echo [ERROR] Build failed.
  exit /b 1
)
echo.
echo Done: %CD%\dist\UnLocked.exe
endlocal
exit /b 0
