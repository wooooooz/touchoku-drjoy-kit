@echo off
rem MonthlySetup.bat - build month config + templates + Mac HTML (pure ASCII launcher)
rem Edit PY below when deploying to the share (portable python path).
setlocal
set "PY=python"
cd /d "%~dp0"
"%PY%" -X utf8 tools\make_month.py --interactive
if errorlevel 1 pause
start "" "%~dp0out"
pause
