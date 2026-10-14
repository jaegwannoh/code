@echo off
rem stop_server.bat - Stop the server started by start_hidden.vbs.
rem Finds the program that uses port 5000 and closes it.
rem (Only that program is closed. Other Python programs are not touched.)

set "FOUND="
for /f "tokens=5" %%a in ('netstat -ano ^| findstr /c:":5000 " ^| findstr LISTENING') do (
  if not defined FOUND (
    set "FOUND=%%a"
    taskkill /F /PID %%a > nul
  )
)

if defined FOUND (
  echo Server stopped. [PID %FOUND%]
) else (
  echo Server is not running.
)
pause
