@echo off
rem kotai_record.bat - update the official duty-change record (R8nenNgatsu\kotai xlsx) for every month folder. 20260929
rem Normally run at the end of the 15:00 sync (dtx14_daily_sync.py). This bat is for manual runs.
set PY=C:\Users\user003\AppData\Local\Programs\Python\Python312\python.exe
cd /d C:\tools\touchoku
"%PY%" -X utf8 -W ignore tools\kotai_record.py %*
