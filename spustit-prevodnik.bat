@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Vytvarim Python prostredi...
    python -m venv .venv
    if errorlevel 1 goto :error
)

".venv\Scripts\python.exe" -c "import fastapi, docx, pypdf, multipart, uvicorn" >nul 2>&1
if errorlevel 1 (
    echo Instaluji potrebne balicky...
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto :error
)

start "Text to Markdown - server" cmd /k ""%~dp0.venv\Scripts\python.exe" -m uvicorn text_to_md.app:app --host 127.0.0.1 --port 8000"
timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:8000"
exit /b 0

:error
echo Aplikaci se nepodarilo spustit. Zkontrolujte instalaci Pythonu a pripojeni k internetu.
pause
exit /b 1
