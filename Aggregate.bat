@echo off
rem Aggregate.bat - collect submissions -> duty JSON + compat xlsx + report (pure ASCII launcher)
rem Edit PY below when deploying to the share (portable python path).
setlocal
set "PY=python"
cd /d "%~dp0"
"%PY%" -X utf8 tools\aggregate.py
if errorlevel 1 pause
pause
