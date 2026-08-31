@echo off
:: Navega para o diretório do projeto
cd /d "c:\Users\muril\OneDrive\FRANQUIAS\master mind\verificar subida de venda pwr"

:: Executa a automação usando o Python do ambiente virtual
.venv\Scripts\python scripts\check_pwr.py

:: Mantém a janela aberta por 5 segundos se executado manualmente para leitura de mensagens de erro rápidas
if "%1"=="manual" (
    timeout /t 5
)
