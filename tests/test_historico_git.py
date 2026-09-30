# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""historico.json versionado no Git (só o arquivo, sem mexer no resto)."""
import os
import shutil
import subprocess
import unittest

from amigo import historico_git

from tests.apoio import TesteIsolado


def git(pasta, *argumentos):
    return subprocess.run(["git", "-C", pasta, *argumentos], capture_output=True, text=True).stdout


@unittest.skipIf(shutil.which("git") is None, "Git não instalado")
class TesteHistoricoGit(TesteIsolado):
    def setUp(self):
        super().setUp()
        self.repo = self._pasta
        git(self.repo, "init", "-q")
        # Repositório sem nome/e-mail: o módulo usa a identidade padrão.
        git(self.repo, "config", "commit.gpgsign", "false")
        self.arquivo = os.path.join(self.repo, "historico.json")
        with open(os.path.join(self.repo, "outro.txt"), "w") as f:
            f.write("alteração em andamento que NÃO pode entrar no commit")
        git(self.repo, "add", "outro.txt")

    def _gravar(self, texto):
        with open(self.arquivo, "w", encoding="utf-8") as f:
            f.write(texto)

    def test_commita_so_o_historico(self):
        self._gravar("[1]")
        self.assertTrue(historico_git.commitar_historico(self.arquivo, "Histórico: CE"))
        self.assertEqual(git(self.repo, "log", "--format=%s").strip(), "Histórico: CE")
        self.assertEqual(git(self.repo, "show", "--name-only", "--format=").split(), ["historico.json"])
        self.assertIn("A  outro.txt", git(self.repo, "status", "--short"))  # continua só preparado

    def test_sem_mudanca_nao_cria_commit(self):
        self._gravar("[1]")
        historico_git.commitar_historico(self.arquivo, "1")
        self.assertFalse(historico_git.commitar_historico(self.arquivo, "2"))
        self.assertEqual(len(git(self.repo, "log", "--format=%s").split()), 1)

    def test_pasta_sem_git_nao_faz_nada(self):
        shutil.rmtree(os.path.join(self.repo, ".git"))
        self._gravar("[1]")
        self.assertFalse(historico_git.commitar_historico(self.arquivo, "x"))

    def test_desligado_nao_dispara(self):
        historico_git.ATIVO = False
        self.assertIsNone(historico_git.guardar_historico_no_git(self.arquivo, "x"))
