@echo off
REM Lance l'application du barometre social.
cd /d "%~dp0"

if not exist config.local.bat (
  copy config.exemple.bat config.local.bat >nul
  echo config.local.bat cree : pensez a changer le code d'acces.
)
call config.local.bat

python -m pip install -q -r requirements.txt
python app.py
pause
