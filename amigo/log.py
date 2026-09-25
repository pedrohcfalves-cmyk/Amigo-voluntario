"""Log amigável: mensagens padronizadas (com hora) no terminal."""
from datetime import datetime

def log_info(mensagem: str):
    """Exibe uma mensagem informativa no terminal, com timestamp."""
    agora = datetime.now().strftime("%H:%M:%S")
    print(f"ℹ️  [{agora}] {mensagem}")


def log_sucesso(mensagem: str):
    """Exibe uma mensagem de sucesso no terminal."""
    agora = datetime.now().strftime("%H:%M:%S")
    print(f"✅ [{agora}] {mensagem}")


def log_erro(mensagem: str):
    """Exibe uma mensagem de erro no terminal."""
    agora = datetime.now().strftime("%H:%M:%S")
    print(f"❌ [{agora}] {mensagem}")


def log_aviso(mensagem: str):
    """Exibe uma mensagem de aviso no terminal."""
    agora = datetime.now().strftime("%H:%M:%S")
    print(f"⚠️  [{agora}] {mensagem}")
