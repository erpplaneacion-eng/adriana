@echo off
title Consolidador de Gastos Mensuales
color 0A
cls
echo ==========================================
echo    CONSOLIDADOR AUTOMATICO DE GASTOS
echo ==========================================
echo.




set /p mes="xxxxxx"
echo.
echo Procesando el mes de %mes%...
echo.
py consolidate_month.py %mes%
echo.
echo ==========================================
echo Presione cualquier tecla para salir.
pause > nul
