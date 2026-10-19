@echo off
rem ============================================================
rem  QDII Quota Dashboard - run one data sync (no web server)
rem  Used by Task Scheduler / manual run.
rem  Safe: no pause / no timeout, always exit with status code.
rem ============================================================
setlocal
cd /d "%~dp0"
set "PY=C:\efunds\Python310\python.exe"
if not exist "%PY%" set "PY=python"
set "LOG=%~dp0sync_task.log"

echo [%date% %time%] ---- sync start ---- >> "%LOG%"
"%PY%" -u "%~dp0sync_data.py" >> "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"
echo [%date% %time%] ---- sync end rc=%RC% ---- >> "%LOG%"

endlocal & exit /b %RC%
