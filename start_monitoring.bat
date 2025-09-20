@echo off
title UGP Reporter - Mode Automatique

echo ============================================================
echo                 UGP REPORTER - MODE AUTOMATIQUE
echo ============================================================
echo.
echo Demarrage du systeme de monitoring intelligent...
echo.

REM Installer les dependances si necessaire
pip install watchdog --quiet 2>nul

REM Creer les dossiers necessaires
mkdir inbox 2>nul
mkdir processed 2>nul
mkdir errors 2>nul
mkdir logs 2>nul

echo Dossiers crees:
echo   - inbox (deposez vos fichiers ici)
echo   - processed (fichiers traites)
echo   - errors (fichiers en erreur)
echo.

REM Lancer le monitoring
python monitoring\auto_processor.py

pause
