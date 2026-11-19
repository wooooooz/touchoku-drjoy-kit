@echo off
rem build_launcher.bat - compile the small launcher (app\launcher\TouchokuLauncher.cs) with the .NET Framework csc.exe
rem that ships with Windows, and pack the unpacked Electron app into dist\touchoku_app.zip + dist\touchoku_app.ver. 20260925
rem Run after "npx electron-builder --win" (needs app\release\win-unpacked).
setlocal
cd /d %~dp0..
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist dist mkdir dist
"%CSC%" /nologo /target:winexe /optimize+ /codepage:65001 /out:dist\touchoku_launcher.exe /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll app\launcher\TouchokuLauncher.cs
if errorlevel 1 ( echo *** build failed: launcher *** & exit /b 1 )
python -X utf8 tools\pack_app.py
if errorlevel 1 ( echo *** pack failed *** & exit /b 1 )
echo done: dist\touchoku_launcher.exe  dist\touchoku_app.zip  dist\touchoku_app.ver
