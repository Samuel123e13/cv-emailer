@echo off
REM Double-click this file to launch CV Emailer.
REM It tries several ways to find Python so it works on a fresh machine.

cd /d "%~dp0"

where py >nul 2>nul && (py app.py & goto :eof)
where python >nul 2>nul && (python app.py & goto :eof)

set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if exist "%PYEXE%" ("%PYEXE%" app.py & goto :eof)

echo Could not find Python. Install it from https://www.python.org/downloads/
echo and tick "Add python.exe to PATH" during setup.
pause
