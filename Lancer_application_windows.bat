@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================================
echo Generateur de demandes d'attestations UDL-SBA
echo ================================================
py -m streamlit run app.py
if errorlevel 1 (
  echo.
  echo Si les dependances ne sont pas installees, executez d'abord :
  echo py -m pip install -r requirements.txt
  pause
)
