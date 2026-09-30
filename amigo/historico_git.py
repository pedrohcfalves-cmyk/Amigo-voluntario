# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
HISTÓRICO NO GIT
================
Depois de cada automação registrada no `historico.json` (a base do
Relatório e das Estatísticas), este módulo grava um commit SÓ desse
arquivo no repositório Git da pasta do programa. Assim o histórico fica
versionado: se o arquivo for apagado ou estragar, dá para recuperar
qualquer versão anterior com o Git.

Regras de segurança (o Git é um "extra" - nunca pode atrapalhar):
  * roda em segundo plano (thread), sem abrir janela preta no Windows;
  * se o Git não estiver instalado, ou a pasta não for um repositório,
    simplesmente não faz nada (avisa 1 vez no painel de mensagens);
  * commita APENAS o historico.json (`git commit -- historico.json`):
    outras alterações que estejam em andamento na pasta não entram;
  * NUNCA envia nada para a internet (não faz push).
"""
import os
import shutil
import subprocess
import threading

from .log import log_aviso

# Desligado nos testes automáticos (ver tests/apoio.py).
ATIVO = True

# Identidade usada só quando o Git do computador não tem nome/e-mail
# configurados (sem isso o commit falharia).
NOME_PADRAO = "Amigo (automático)"
EMAIL_PADRAO = "amigo@localhost"

TEMPO_LIMITE_SEGUNDOS = 30

_ja_avisou = False
_trava = threading.Lock()  # 2 commits ao mesmo tempo brigariam pelo index.lock


def _rodar_git(pasta, *argumentos):
    opcoes = {}
    if os.name == "nt":
        opcoes["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.run(
        ["git", "-C", pasta, *argumentos],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=TEMPO_LIMITE_SEGUNDOS, **opcoes,
    )


def _avisar_uma_vez(mensagem):
    global _ja_avisou
    if not _ja_avisou:
        _ja_avisou = True
        log_aviso(mensagem)


def commitar_historico(caminho_arquivo, mensagem):
    """Faz o commit (sincrono). Devolve True se gravou um commit novo."""
    if shutil.which("git") is None:
        _avisar_uma_vez("Histórico não foi guardado no Git: o Git não está instalado neste computador "
                        "(o historico.json continua salvo normalmente).")
        return False

    pasta, nome = os.path.split(os.path.abspath(caminho_arquivo))
    with _trava:
        try:
            if _rodar_git(pasta, "rev-parse", "--is-inside-work-tree").returncode != 0:
                return False  # pasta sem Git: nada a fazer, sem aviso

            identidade = []
            if not _rodar_git(pasta, "config", "user.email").stdout.strip():
                identidade = ["-c", f"user.name={NOME_PADRAO}", "-c", f"user.email={EMAIL_PADRAO}"]

            adicionar = _rodar_git(pasta, "add", "--", nome)
            if adicionar.returncode != 0:
                _avisar_uma_vez(f"Histórico não foi guardado no Git: {adicionar.stderr.strip()[:200]}")
                return False

            # Nada mudou desde o último commit? Então não há o que gravar.
            if _rodar_git(pasta, "diff", "--cached", "--quiet", "--", nome).returncode == 0:
                return False

            commit = _rodar_git(pasta, *identidade, "commit", "-q", "-m", mensagem, "--", nome)
            if commit.returncode != 0:
                _avisar_uma_vez(f"Histórico não foi guardado no Git: {commit.stderr.strip()[:200]}")
                return False
            return True
        except (OSError, subprocess.SubprocessError) as erro:
            _avisar_uma_vez(f"Histórico não foi guardado no Git: {erro}")
            return False


def guardar_historico_no_git(caminho_arquivo, mensagem):
    """Versão em segundo plano, usada depois de cada automação."""
    if not ATIVO:
        return None
    tarefa = threading.Thread(
        target=commitar_historico, args=(caminho_arquivo, mensagem),
        name="historico-git", daemon=True,
    )
    tarefa.start()
    return tarefa
