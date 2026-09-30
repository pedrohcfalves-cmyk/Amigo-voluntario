@echo off
rem ==============================================================================
rem Amigo - gera o aplicativo executavel (Amigo.exe) com PyInstaller.
rem ==============================================================================
rem Script sem nenhuma pergunta/pausa no meio (roda do inicio ao fim sozinho).
rem Grava o resultado em build_status.txt (OK ou FALHOU) e o log completo em
rem build.log, para conferencia mesmo depois que a janela fechar.
rem ==============================================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"

del /q build_status.txt >nul 2>nul
del /q py_probe.txt >nul 2>nul
echo Iniciando build em %date% %time% > build.log

echo ==============================================================
echo Amigo - construindo o aplicativo executavel (Amigo.exe)
echo ==============================================================
echo.

rem Descobre qual comando roda o Python de verdade nesta sessao, SEM
rem usar pipes (evita travamentos) - so redireciona a saida para um
rem arquivo e confere o conteudo. Em alguns computadores, "python" no
rem PATH aponta para o "stub" da Microsoft Store (nao o Python de
rem verdade) - nesse caso ele so imprime uma mensagem sobre instalar
rem pela Store, sem rodar nada - por isso testamos rodando de verdade.
echo Procurando instalacao do Python... >>build.log
set "PY_CMD="

python -c "print(1)" >py_probe.txt 2>&1
findstr /X "1" py_probe.txt >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"

if not defined PY_CMD (
    py -3 -c "print(1)" >py_probe.txt 2>&1
    findstr /X "1" py_probe.txt >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py -3"
)

if not defined PY_CMD (
    py -c "print(1)" >py_probe.txt 2>&1
    findstr /X "1" py_probe.txt >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py"
)

del /q py_probe.txt >nul 2>nul

if not defined PY_CMD (
    echo O Python nao foi encontrado - ou so aponta para o atalho da Microsoft Store - neste computador. 1>>build.log
    echo O Python nao foi encontrado neste computador.
    echo FALHOU: Python nao encontrado > build_status.txt
    goto fim
)
echo Usando Python: !PY_CMD! >>build.log

echo [1/4] Instalando PyInstaller...
!PY_CMD! -m pip install --upgrade pyinstaller >>build.log 2>&1
if errorlevel 1 (
    echo Falha ao instalar o PyInstaller. Veja build.log.
    echo FALHOU: pip install pyinstaller > build_status.txt
    goto fim
)

echo [2/4] Instalando dependencias do programa (pywin32, playwright)...
!PY_CMD! -m pip install -r requirements.txt >>build.log 2>&1
if errorlevel 1 (
    echo Falha ao instalar as dependencias do programa. Veja build.log.
    echo FALHOU: pip install -r requirements.txt > build_status.txt
    goto fim
)

echo [3/4] Gerando o executavel (pode levar alguns minutos)...
!PY_CMD! -m PyInstaller --noconfirm --onefile --windowed --name Amigo --collect-all playwright main_gui.py >>build.log 2>&1
if errorlevel 1 (
    echo Falha ao gerar o executavel. Veja build.log.
    echo FALHOU: pyinstaller > build_status.txt
    goto fim
)

echo [4/4] Copiando Amigo.exe para a pasta principal...
if exist "Amigo.exe" (
    del /q "Amigo.exe.antigo" >nul 2>nul
    ren "Amigo.exe" "Amigo.exe.antigo" >nul 2>>build.log
)

set TENTATIVAS_COPIA=0
:tentar_copiar_exe
set /a TENTATIVAS_COPIA+=1
copy /y "dist\Amigo.exe" "Amigo.exe" >nul 2>>build.log
if not errorlevel 1 goto copiou_exe
if !TENTATIVAS_COPIA! GEQ 8 (
    echo Falha ao copiar Amigo.exe para a pasta principal. Veja build.log.
    echo FALHOU: copiar exe > build_status.txt
    goto fim
)
echo   Tentando de novo em instantes... ^(tentativa !TENTATIVAS_COPIA!^)
timeout /t 3 /nobreak >nul
goto tentar_copiar_exe
:copiou_exe
del /q "Amigo.exe.antigo" >nul 2>nul

echo.
echo ==============================================================
echo CONCLUIDO: Amigo.exe criado com sucesso nesta pasta.
echo ==============================================================
echo OK > build_status.txt

:fim
echo.
echo (Esta janela fecha sozinha em 30 segundos - build_status.txt e build.log ficam salvos na pasta.)
timeout /t 30
