@echo off
:: Navega para a pasta do sistema
cd /d "%~dp0sistema"

:: Executa a automação local
.venv\Scripts\python check_pwr.py --local

echo.
echo ============================================================
echo Processamento concluido! 
echo Pressione qualquer tecla para fechar.
pause > nul
