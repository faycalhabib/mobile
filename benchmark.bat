@echo off
title UGP Reporter - Benchmark Performance

echo ============================================================
echo                  BENCHMARK DE PERFORMANCE
echo ============================================================
echo.
echo Ce test va comparer les performances entre:
echo   1. Mode CLASSIQUE (Win32COM) - Stable mais lent
echo   2. Mode OPTIMISE (openpyxl) - Rapide et efficace
echo.
echo Appuyez sur une touche pour demarrer...
pause >nul

REM Lancer le benchmark
python test_performance.py

echo.
echo ============================================================
echo Pour basculer entre les modes:
echo   python toggle_fast_mode.py
echo.
echo Mode actuel:
python toggle_fast_mode.py status
echo ============================================================
pause
