@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" goto create_venv
rem A .venv copied from another computer points at that computer's Python; rebuild it.
".venv\Scripts\python.exe" -c "import sys" >nul 2>&1
if errorlevel 1 (
    echo Prostredi .venv je z jineho pocitace, pripravim nove...
    rmdir /s /q ".venv"
    goto create_venv
)

goto check_dependencies

:create_venv
echo Vytvarim virtualni prostredi .venv...
set "PYTHON_CMD="
where py >nul 2>&1
if not errorlevel 1 (
    py -3 --version >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=py -3"
)
if defined PYTHON_CMD goto run_venv_creation
where python >nul 2>&1
if not errorlevel 1 (
    python --version >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=python"
)
if defined PYTHON_CMD goto run_venv_creation

echo.
echo Python se ted jednou nainstaluje. Chvili pockej, muze to trvat nekolik minut.
echo Pokud se objevi okno s dotazem na povoleni, klikni na Ano.
echo.
where winget >nul 2>&1
if errorlevel 1 goto python_missing
winget install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements

rem PATH se v tomto okne neobnovil, proto hledame Python primo.
set "PY_LOCAL=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
set "PY_LAUNCHER=%LOCALAPPDATA%\Programs\Python\Launcher\py.exe"
if exist "%PY_LOCAL%" set PYTHON_CMD="%PY_LOCAL%"
if defined PYTHON_CMD goto run_venv_creation
py -3 --version >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3"
if defined PYTHON_CMD goto run_venv_creation
if exist "%PY_LAUNCHER%" (
    "%PY_LAUNCHER%" -3 --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD="%PY_LAUNCHER%" -3
)
if not defined PYTHON_CMD goto python_missing

:run_venv_creation
%PYTHON_CMD% -m venv .venv
if errorlevel 1 goto venv_failed
if not exist ".venv\Scripts\python.exe" goto venv_failed

echo.
echo Virtualni prostredi je pripravene.

:check_dependencies
.venv\Scripts\python.exe -c "import docx; from PIL import Image" >nul 2>&1
if errorlevel 1 goto install_dependencies

goto run_app

:install_dependencies
echo Instaluji zavislosti z requirements.txt...
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto install_failed

goto run_app

:run_app
echo Spoustim Klikatko. Toto okno nech otevrene, dokud pracujes. Az skoncis, zavri ho krizkem.
echo.
.venv\Scripts\python.exe main.py %*
if errorlevel 1 goto app_failed
exit /b 0

:python_missing
echo CHYBA: Python se nepodarilo nainstalovat.
echo Zavolej prosim Barbore, ona to pripravi.
goto failed

:venv_failed
echo CHYBA: Nepodarilo se vytvorit virtualni prostredi .venv.
goto failed

:install_failed
echo CHYBA: Instalace zavislosti selhala. Zkontroluj internet a spust spustit.bat znovu. Kdyz to nepujde, zavolej Barbore.
goto failed

:app_failed
echo CHYBA: Aplikace skoncila s chybou. Vyfot tohle okno a posli to Barbore.
goto failed

:failed
echo.
pause
exit /b 1
