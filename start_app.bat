@echo off
setlocal
cd /d "%~dp0"
set "APP_VERSION=2.0"
set "REQ_STAMP=.venv\requirements-v2.0-installed.txt"
if not exist .venv\Scripts\python.exe (
  echo Creating a local Python virtual environment...
  py -3 -m venv .venv
  if errorlevel 1 goto error
)
call .venv\Scripts\activate.bat
if not exist "%REQ_STAMP%" (
  echo Installing or updating the app packages for version %APP_VERSION%. An internet connection is required for this step.
  python -m pip install --upgrade -r requirements.txt
  if errorlevel 1 goto error
  type nul > "%REQ_STAMP%"
)
python -m streamlit run app.py
if errorlevel 1 goto error
exit /b 0
:error
echo.
echo App could not start. Check the error above and confirm Python 3.10 or newer is installed.
pause
exit /b 1
