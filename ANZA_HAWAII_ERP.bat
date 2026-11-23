@echo off
setlocal
cd /d "%~dp0"
set APP_ENV=development
set COOKIE_SECURE=0
if not exist .venv (
  echo Inatengeneza mazingira ya Python...
  py -3 -m venv .venv
  if errorlevel 1 goto python_error
)
call .venv\Scripts\activate.bat
 echo Inaweka mahitaji ya mfumo...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto install_error
 echo.
echo HAWAII ERP inaanza. Fungua http://127.0.0.1:5000 kwenye Chrome.
echo Mara ya kwanza utapelekwa kwenye ukurasa wa kutengeneza akaunti ya Admin.
python app.py
pause
exit /b
:python_error
echo Imeshindikana kuanzisha Python. Hakikisha Python 3 imewekwa na chaguo la PATH limewashwa.
pause
exit /b 1
:install_error
echo Imeshindikana kusakinisha packages. Angalia internet kisha jaribu tena.
pause
exit /b 1
