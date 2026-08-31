@echo off
:: Navega para a pasta local (sistema/)
cd /d "%~dp0"

:: Executa a automação usando o Python do ambiente virtual
.venv\Scripts\python check_pwr.py

:: Mantém a janela aberta por 5 segundos se executado manualmente para leitura de mensagens de erro rápidas
if "%1"=="manual" (
    timeout /t 5
)
