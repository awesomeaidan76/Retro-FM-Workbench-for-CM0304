@echo off
setlocal
cd /d "%~dp0"
py -3 retrofm\app.py
if errorlevel 1 (
  echo.
  echo Retro FM Workbench failed to start.
  echo Make sure Python 3.11+ is installed with Tcl/Tk.
  pause
)
