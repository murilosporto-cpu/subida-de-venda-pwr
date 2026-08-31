@echo off
:: Navega para o diretório do projeto
cd /d "%~dp0"

:: Executa a automação local
.venv\Scripts\python check_pwr.py --local

echo.
echo ============================================================
echo Processamento concluido! 
echo Pressione qualquer tecla para fechar.
pause > nul
