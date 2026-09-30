@echo off
rem ==============================================================================
rem Amigo - Automacoes SIGEF (atalho de duplo clique)
rem ==============================================================================
rem Abre a interface grafica do sistema. Na primeira vez, o proprio aplicativo
rem baixa/instala automaticamente o que faltar (pacotes Python, navegador do
rem Playwright) antes de liberar as automacoes.
rem ==============================================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo O Python nao foi encontrado neste computador.
    echo Instale o Python 3.10 ou mais novo e tente novamente:
    echo https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

where pythonw >nul 2>nul
if errorlevel 1 (
    python main_gui.py
) else (
    start "Amigo" pythonw main_gui.py
)
