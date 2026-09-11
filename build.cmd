@echo off
rem Builds WinCleaner.exe with the C# compiler that ships with Windows (.NET Framework 4.x).
setlocal
cd /d "%~dp0"
set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo csc.exe not found - .NET Framework 4.x is required.
  exit /b 1
)
if not exist "src\app.ico" powershell -NoProfile -ExecutionPolicy Bypass -File "src\make-icon.ps1"
"%CSC%" /nologo /codepage:65001 /target:winexe /optimize+ /platform:anycpu /out:WinCleaner.exe ^
  /win32manifest:src\app.manifest /win32icon:src\app.ico ^
  /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.ServiceProcess.dll ^
  src\WinCleaner.cs
if errorlevel 1 exit /b 1
echo Built: %~dp0WinCleaner.exe
