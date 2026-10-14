@echo off
rem ======================================================================
rem  Daily work log server (Python 2.7) - double-click to run.
rem  (Comments are in English on purpose: Korean text in .bat files can
rem   break on English Windows.)
rem ======================================================================

rem ---- [1] Python path ---------------------------------------------------
rem  Leave empty to search common Anaconda folders automatically.
rem  If it fails, put the full path of python.exe here, for example:
rem    set "MY_PYTHON=C:\Anaconda2\python.exe"
rem  (To find it: open Anaconda Prompt and type:  where python )
set "MY_PYTHON="
rem ------------------------------------------------------------------------

set "PYTHON=%MY_PYTHON%"
if not defined PYTHON (
  for %%P in (
    "C:\Anaconda2\python.exe"
    "C:\ProgramData\Anaconda2\python.exe"
    "%USERPROFILE%\Anaconda2\python.exe"
    "%LOCALAPPDATA%\Continuum\anaconda2\python.exe"
    "C:\Anaconda\python.exe"
    "C:\ProgramData\Anaconda\python.exe"
    "%USERPROFILE%\Anaconda\python.exe"
  ) do (
    if not defined PYTHON if exist "%%~P" set "PYTHON=%%~P"
  )
)
if not defined PYTHON set "PYTHON=python"

rem PYTHONIOENCODING : print Korean text as UTF-8 instead of crashing.
set PYTHONIOENCODING=utf-8

cd /d "%~dp0"

rem ---- Already running? (something is using port 5000 = PORT in app.py) --
netstat -ano | findstr /c:":5000 " | findstr LISTENING > nul
if not errorlevel 1 (
  if /i "%~1"=="hidden" exit /b
  echo The server is already running. Open http://localhost:5000
  echo To stop it, run stop_server.bat
  pause
  exit /b
)

rem ---- Hidden mode (started by start_hidden.vbs) -------------------------
rem  No window: app.py --log writes all messages to server.log itself.
rem  The previous log is kept as server_old.log. When server.log grows over
rem  5 MB it is renamed to server_old.log and a new one is started.
rem  server_boot.log only catches errors before app.py starts (normally empty).
if /i "%~1"=="hidden" (
  "%PYTHON%" app.py --log > server_boot.log 2>&1
  exit /b
)

echo Python: %PYTHON%
"%PYTHON%" --version
"%PYTHON%" app.py
pause
