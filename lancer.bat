@echo off
REM Double-click to start the application (creates a local Python environment on first run).
cd /d "%~dp0"
if not exist .venv (
    echo Premiere installation, patiente quelques secondes...
    py -3 -m venv .venv 2>nul || python -m venv .venv
    if errorlevel 1 (
        echo Python est introuvable. Installe-le depuis python.org en cochant "Add python.exe to PATH".
        pause
        exit /b 1
    )
)
.venv\Scripts\python.exe -m pip install -q --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
    echo L installation des dependances a echoue : verifie ta connexion internet.
    pause
    exit /b 1
)
start "" .venv\Scripts\pythonw.exe -m app
