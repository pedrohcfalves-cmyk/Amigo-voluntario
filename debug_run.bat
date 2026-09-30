@echo off
cd /d "%~dp0"
echo Iniciando main_gui.py (modo script, com console para depuracao)... > debug_log.txt
python main_gui.py >> debug_log.txt 2>&1
echo. >> debug_log.txt
echo CODIGO DE SAIDA: %errorlevel% >> debug_log.txt
echo Concluido. Log em debug_log.txt
timeout /t 10
