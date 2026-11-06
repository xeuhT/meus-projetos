@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ======================================================================
echo   ROBO DE EXTRACAO DE EXTRATO DE COMPRAS - GOODMANAGER (TICKET PLUS)
echo ======================================================================
python robo_extrato_ticket.py %*
echo.
pause
