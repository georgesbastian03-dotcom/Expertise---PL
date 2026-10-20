@echo off
setlocal
cd /d "%~dp0"

rem Recherche de Python. "python" peut n'etre qu'un raccourci du Microsoft
rem Store (dossier WindowsApps) qui ne lance rien : on l'ignore et on essaie
rem d'abord le lanceur "py", puis les emplacements d'installation habituels.
set "PY="
py -3 -c "import tkinter" >nul 2>nul && set "PY=py -3"
if not defined PY (
    for /f "delims=" %%P in ('where python 2^>nul ^| findstr /v /i "WindowsApps"') do (
        if not defined PY set "PY="%%P""
    )
)
if not defined PY (
    for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*" "%ProgramFiles%\Python3*") do (
        if exist "%%~D\python.exe" set "PY="%%~D\python.exe""
    )
)
if not defined PY (
    echo Python n'a pas ete trouve sur cet ordinateur.
    echo Installez-le depuis https://www.python.org/downloads/ puis relancez ce fichier.
    echo Important : lors de l'installation, cochez la case "Add python.exe to PATH".
    pause
    exit /b 1
)

rem Dependances : reinstallees si elles manquent pour ce Python-la (par
rem exemple apres une reinstallation ou un changement de version de Python).
%PY% -c "import openpyxl, docx, fitz, PIL" >nul 2>nul
if errorlevel 1 (
    echo Installation des dependances necessaires ^(une seule fois^)...
    %PY% -m pip install --quiet --upgrade pip
    %PY% -m pip install --quiet -r requirements.txt
    if errorlevel 1 (
        echo.
        echo L'installation des dependances a echoue. Verifiez votre connexion internet.
        pause
        exit /b 1
    )
)
echo ok > .deps_ok

%PY% main.py
if errorlevel 1 pause
