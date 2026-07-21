@echo off
title CRM Prospecção - Servidor Local
echo ========================================================
echo               Iniciando CRM Prospecção...
echo ========================================================
echo.
python app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] Ocorreu um erro ao iniciar a aplicação.
    pause
)
