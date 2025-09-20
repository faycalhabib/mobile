@echo off
title Arrêt du Monitoring UGP

echo ============================================================
echo          ARRÊT DU MONITORING UGP REPORTER
echo ============================================================
echo.

REM Arrêter tous les processus Python
echo Arrêt des processus Python en cours...
taskkill /IM python.exe /F 2>nul

if %ERRORLEVEL% == 0 (
    echo.
    echo ✓ Monitoring arrêté avec succès
) else (
    echo.
    echo ℹ Aucun processus Python en cours
)

echo.
echo ============================================================
echo Pour redémarrer le monitoring:
echo   python monitoring/auto_processor.py
echo ============================================================
pause
