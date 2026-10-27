@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" goto create_venv

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
echo Spoustim Klikatko. Aplikaci ukoncite klavesami Ctrl+C.
echo.
.venv\Scripts\python.exe main.py %*
if errorlevel 1 goto app_failed
exit /b 0

:python_missing
echo CHYBA: Nebyl nalezen Python 3. Nainstalujte Python 3.10 nebo novejsi a zkuste to znovu.
goto failed

:venv_failed
echo CHYBA: Nepodarilo se vytvorit virtualni prostredi .venv.
goto failed

:install_failed
echo CHYBA: Instalace zavislosti selhala. Zkontrolujte pripojeni k internetu a zkuste spustit skript znovu.
goto failed

:app_failed
echo CHYBA: Aplikace skoncila s chybou. Zkontrolujte vypis vyse.
goto failed

:failed
echo.
pause
exit /b 1
