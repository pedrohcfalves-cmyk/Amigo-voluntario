# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
"BAIXA TUDO NECESSÁRIO"
========================
Verifica e instala, automaticamente, tudo que falta para o sistema rodar:
os pacotes Python (pywin32, playwright) e o navegador Chromium usado pelo
Playwright para automatizar o SIGEF.

Usado pela interface gráfica (`amigo/gui.py`) antes de liberar os botões
de automação - assim o usuário só precisa abrir o aplicativo (ou o
Amigo.bat) e esperar, sem rodar nenhum comando manualmente.

Roda em dois "modos" diferentes:
    - Como script Python comum (`python main_gui.py` / Amigo.bat): pode
      instalar pacotes Python que faltarem via pip, e chama o instalador
      do Chromium via subprocesso (`python -m playwright install`).
    - Empacotado como .exe (PyInstaller): os pacotes Python já vêm
      embutidos no .exe (nada para instalar), só falta o navegador
      Chromium - chamado EM PROCESSO (`playwright.__main__`), porque
      dentro do .exe não existe um "python -m" de verdade para chamar via
      subprocesso.
"""
import importlib.util
import os
import subprocess
import sys
from typing import Callable, List

from .config_io import BASE_DIR

CAMINHO_REQUIREMENTS = os.path.join(BASE_DIR, "requirements.txt")


def _rodando_empacotado() -> bool:
    """True quando o programa está rodando como .exe gerado pelo
    PyInstaller, e não como script Python comum."""
    return bool(getattr(sys, "frozen", False))


def _pacote_instalado(nome_modulo: str) -> bool:
    try:
        return importlib.util.find_spec(nome_modulo) is not None
    except (ImportError, ValueError):
        return False


def _pasta_cache_playwright() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ms-playwright")


def _chromium_instalado() -> bool:
    """
    Verifica se o Chromium do Playwright já foi baixado, olhando
    diretamente a pasta de cache (mais rápido e não depende de rede) -
    evita reinstalar a cada vez que o aplicativo é aberto.
    """
    pasta = _pasta_cache_playwright()
    if not os.path.isdir(pasta):
        return False
    try:
        return any(nome.startswith("chromium-") for nome in os.listdir(pasta))
    except OSError:
        return False


def verificar_pendencias() -> List[str]:
    """Devolve uma lista de descrições curtas do que falta instalar (lista
    vazia se tudo já estiver pronto - nada a fazer)."""
    pendencias = []
    if not _rodando_empacotado():
        if not _pacote_instalado("win32com"):
            pendencias.append("pacote Python 'pywin32'")
        if not _pacote_instalado("playwright"):
            pendencias.append("pacote Python 'playwright'")
    if not _chromium_instalado():
        pendencias.append("navegador Chromium do Playwright (download)")
    return pendencias


def instalar_dependencias(logar: Callable[[str], None]) -> bool:
    """
    Instala tudo que `verificar_pendencias()` apontar como faltante,
    chamando `logar(mensagem)` a cada passo para a interface mostrar o
    progresso. Devolve True se tudo terminou OK (ou já não havia nada para
    instalar).
    """
    ok = True

    if not _rodando_empacotado() and (
        not _pacote_instalado("win32com") or not _pacote_instalado("playwright")
    ):
        logar("Instalando pacotes Python (pywin32, playwright)...")
        try:
            resultado = subprocess.run(
                [sys.executable, "-m", "pip", "install", "-r", CAMINHO_REQUIREMENTS],
                capture_output=True, text=True, timeout=600,
            )
            for linha in (resultado.stdout or "").splitlines():
                logar(linha)
            if resultado.returncode != 0:
                ok = False
                logar("Falha ao instalar pacotes Python:")
                for linha in (resultado.stderr or "").splitlines():
                    logar(linha)
            else:
                logar("Pacotes Python instalados.")
        except Exception as erro:
            ok = False
            logar(f"Falha ao instalar pacotes Python: {erro}")

    if ok and not _chromium_instalado():
        logar("Baixando o navegador Chromium do Playwright (só na primeira vez; pode levar alguns minutos)...")
        try:
            if _rodando_empacotado():
                # Dentro do .exe, sys.executable é o próprio Amigo.exe (não
                # um python.exe de verdade) - por isso é preciso chamar o
                # instalador do Playwright EM PROCESSO, não via subprocesso.
                from playwright.__main__ import main as playwright_main
                argv_original = sys.argv
                sys.argv = ["playwright", "install", "chromium"]
                try:
                    playwright_main()
                except SystemExit as saida:
                    if saida.code not in (0, None):
                        ok = False
                finally:
                    sys.argv = argv_original
            else:
                resultado = subprocess.run(
                    [sys.executable, "-m", "playwright", "install", "chromium"],
                    capture_output=True, text=True, timeout=900,
                )
                for linha in (resultado.stdout or "").splitlines():
                    logar(linha)
                if resultado.returncode != 0:
                    ok = False
                    for linha in (resultado.stderr or "").splitlines():
                        logar(linha)
            if ok:
                logar("Navegador Chromium instalado.")
        except Exception as erro:
            ok = False
            logar(f"Falha ao instalar o navegador Chromium: {erro}")

    return ok
