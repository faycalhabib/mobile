@echo off
echo ====================================
echo    UGP Reporter - Demarrage
echo ====================================
echo.

python main.py

if errorlevel 1 (
    echo.
    echo ERREUR: L'application a rencontre une erreur.
    echo Verifiez les logs dans le dossier logs/
    pause
)
