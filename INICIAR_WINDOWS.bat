@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Creando entorno virtual...
  py -m venv .venv
  if errorlevel 1 goto :error
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

if not exist ".streamlit\secrets.toml" (
  copy ".streamlit\secrets.toml.example" ".streamlit\secrets.toml" >nul
  echo.
  echo Se creo .streamlit\secrets.toml con valores de ejemplo.
  echo Editalo con tu GEMINI_API_KEY, client_id y client_secret antes de continuar.
  pause
  exit /b 0
)

python -m streamlit run app.py
exit /b 0

:error
echo.
echo No fue posible preparar la aplicacion. Revisa el mensaje anterior.
pause
exit /b 1
