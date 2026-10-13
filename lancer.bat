@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python n'a pas ete trouve sur cet ordinateur.
    echo Installez-le depuis https://www.python.org/downloads/ puis relancez ce fichier.
    echo Important : lors de l'installation, cochez la case "Add Python to PATH".
    pause
    exit /b 1
)

if not exist ".deps_ok" (
    echo Installation des dependances necessaires ^(une seule fois^)...
    python -m pip install --quiet --upgrade pip
    python -m pip install --quiet -r requirements.txt
    if errorlevel 1 (
        echo.
        echo L'installation des dependances a echoue. Verifiez votre connexion internet.
        pause
        exit /b 1
    )
    echo ok > .deps_ok
)

python main.py
if errorlevel 1 pause
