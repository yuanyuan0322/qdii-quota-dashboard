@echo off
rem ============================================================
rem  QDII Quota Dashboard - start web server (http://localhost:5001)
rem  Safe for Task Scheduler: no pause / no timeout, always exit.
rem ============================================================
setlocal
cd /d "%~dp0"
set "PY=C:\efunds\Python310\python.exe"
if not exist "%PY%" set "PY=python"
set "RC=0"

netstat -ano | findstr /R /C:"LISTENING" | findstr ":5001 " >nul 2>&1
if not errorlevel 1 (
    echo [%date% %time%] port 5001 already listening, skip >> "%~dp0keepalive.log"
    goto :done
)

echo [%date% %time%] starting QDII dashboard server >> "%~dp0keepalive.log"
start "QDII Dashboard" /min "%PY%" "%~dp0app.py"

:done
endlocal & exit /b %RC%
