@echo off
chcp 65001 > nul
title Conferência e Carga de Seguro de Vida - Cencosud Brasil
echo ===============================================================================
echo      AUTOMAÇÃO DA CONFERÊNCIA DE SEGURO DE VIDA - CENCOSUD BRASIL
echo ===============================================================================
echo.
echo Processando conferência e gerando cargas por bandeira...
echo.
python "c:\Users\matcaval\OneDrive - Cencosud\Documentos\Python Antigravity Agent\conferencia_seguro_vida\main.py"
echo.
echo ===============================================================================
echo Processamento concluído! Pressione qualquer tecla para fechar...
echo ===============================================================================
pause > nul
