@echo off
rem Roda todos os testes automaticos do Amigo (nao abre o SIGEF nem o Excel).
cd /d "%~dp0"
python -m unittest discover -s tests -t . -v
echo.
pause
