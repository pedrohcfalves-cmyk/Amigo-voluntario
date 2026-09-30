# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
Log amigável: mensagens padronizadas (com hora) no terminal.

Além de imprimir no terminal (comportamento original, usado por
`main.py`), cada função também repassa a mensagem para "ouvintes"
registrados via `adicionar_ouvinte()` - usado pela interface gráfica
(`amigo/gui.py`) para mostrar as mesmas mensagens numa caixa de texto,
sem duplicar nem alterar nenhuma chamada existente a log_info/log_sucesso/
log_erro/log_aviso em todo o resto do sistema.
"""
from datetime import datetime

# Funções chamadas, na ordem em que foram registradas, a cada mensagem de
# log - assinatura: ouvinte(nivel: str, mensagem: str). Uma exceção dentro
# de um ouvinte NUNCA pode derrubar a automação: é sempre engolida aqui.
_ouvintes = []


def _imprimir(texto: str) -> None:
    """
    `print()` protegido: quando empacotado com PyInstaller no modo
    `--windowed` (sem console), `sys.stdout`/`sys.stderr` ficam `None` no
    Windows, e um `print()` comum levantaria `AttributeError` e derrubaria
    a automação em andamento. No modo terminal (`python main.py`) o
    comportamento continua idêntico - imprime normalmente.
    """
    try:
        print(texto)
    except Exception:
        pass


def adicionar_ouvinte(ouvinte) -> None:
    """Registra uma função para ser chamada a cada log_info/sucesso/erro/aviso."""
    if ouvinte not in _ouvintes:
        _ouvintes.append(ouvinte)


def remover_ouvinte(ouvinte) -> None:
    """Remove um ouvinte registrado com `adicionar_ouvinte()` (se existir)."""
    if ouvinte in _ouvintes:
        _ouvintes.remove(ouvinte)


def _notificar(nivel: str, mensagem: str) -> None:
    for ouvinte in list(_ouvintes):
        try:
            ouvinte(nivel, mensagem)
        except Exception:
            pass


def log_info(mensagem: str):
    """Exibe uma mensagem informativa no terminal, com timestamp."""
    agora = datetime.now().strftime("%H:%M:%S")
    _imprimir(f"ℹ️  [{agora}] {mensagem}")
    _notificar("info", mensagem)


def log_sucesso(mensagem: str):
    """Exibe uma mensagem de sucesso no terminal."""
    agora = datetime.now().strftime("%H:%M:%S")
    _imprimir(f"✅ [{agora}] {mensagem}")
    _notificar("sucesso", mensagem)


def log_erro(mensagem: str):
    """Exibe uma mensagem de erro no terminal."""
    agora = datetime.now().strftime("%H:%M:%S")
    _imprimir(f"❌ [{agora}] {mensagem}")
    _notificar("erro", mensagem)


def log_aviso(mensagem: str):
    """Exibe uma mensagem de aviso no terminal."""
    agora = datetime.now().strftime("%H:%M:%S")
    _imprimir(f"⚠️  [{agora}] {mensagem}")
    _notificar("aviso", mensagem)
