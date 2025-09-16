@echo off
echo ====================================
echo Installation de UGP Reporter
echo ====================================
echo.

echo [1/3] Verification de Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERREUR: Python n'est pas installe!
    echo Veuillez installer Python depuis https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [2/3] Installation des dependances...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo [3/3] Creation des dossiers...
if not exist "outputs" mkdir outputs
if not exist "logs" mkdir logs
if not exist "config" mkdir config

echo.
echo ====================================
echo Installation terminee avec succes!
echo ====================================
echo.
echo Pour lancer l'application: run.bat
echo.
pause
