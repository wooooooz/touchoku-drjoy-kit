@echo off
rem dtx14_daily_sync.bat - run the 15:00 sync (shared master -> DTX14 -> public PDF). 20260929
rem Registered in Task Scheduler as "Touchoku_DTX14_sync" (daily 15:00, the administrator's PC).
rem Python errors that happen before logging starts (import errors etc.) go to a local file.
rem Exit code 1 = some month/stage failed (an alert mail was sent by the script).
set PY=C:\Users\user003\AppData\Local\Programs\Python\Python312\python.exe
set LOGDIR=%LOCALAPPDATA%\Touchoku\log
if not exist "%LOGDIR%" mkdir "%LOGDIR%"
cd /d C:\tools\touchoku
"%PY%" -X utf8 -W ignore tools\dtx14_daily_sync.py %* 2>> "%LOGDIR%\dtx14_daily_sync.err"
exit /b %ERRORLEVEL%
