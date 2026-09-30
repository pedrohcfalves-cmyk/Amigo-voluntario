# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""Ajudantes comuns dos testes: pasta temporária no lugar de config.json /
historico.json, e as colunas sempre no padrão de fábrica."""
import os
import shutil
import tempfile
import unittest

from unittest import mock

from amigo import colunas, config_io, execucao, historico_git, log, relatorio


class TesteIsolado(unittest.TestCase):
    """Base: cada teste grava config/histórico numa pasta temporária e
    começa com as colunas de fábrica e a simulação desligada."""

    def setUp(self):
        self._pasta = tempfile.mkdtemp(prefix="amigo_teste_")
        self._config_original = config_io.CONFIG_PATH
        self._historico_original = relatorio.HISTORICO_PATH
        config_io.CONFIG_PATH = os.path.join(self._pasta, "config.json")
        relatorio.HISTORICO_PATH = os.path.join(self._pasta, "historico.json")
        colunas.aplicar_colunas(dict(colunas.COLUNAS_PADRAO))
        execucao.finalizar_execucao()
        historico_git.ATIVO = False  # nenhum teste faz commit de verdade no Git
        # Mensagens das automações não poluem a saída dos testes.
        self._silencio = mock.patch.object(log, "_imprimir", lambda _texto: None)
        self._silencio.start()

    def tearDown(self):
        self._silencio.stop()
        config_io.CONFIG_PATH = self._config_original
        relatorio.HISTORICO_PATH = self._historico_original
        colunas.aplicar_colunas(dict(colunas.COLUNAS_PADRAO))
        execucao.finalizar_execucao()
        shutil.rmtree(self._pasta, ignore_errors=True)
