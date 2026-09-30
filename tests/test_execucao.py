# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""Controle de execução: simulação, parar e progresso (itens 4 e 5)."""
from amigo import execucao
from amigo.execucao import linha_em_execucao
from amigo.lancamento_manual import PlanilhaVirtual
from amigo.log import log_aviso

from tests.apoio import TesteIsolado


class TesteExecucao(TesteIsolado):
    def test_marca_de_simulado(self):
        self.assertTrue(execucao.eh_simulado("SIMULADO ✓ CE conferida"))
        self.assertTrue(execucao.eh_simulado("  simulado ✗ x"))
        self.assertFalse(execucao.eh_simulado("2027CE000111"))
        self.assertFalse(execucao.eh_simulado(None))

    def test_parar(self):
        execucao.iniciar_execucao()
        self.assertFalse(execucao.parar_antes_da_linha(5))
        execucao.pedir_parada()
        self.assertTrue(execucao.parar_antes_da_linha(5))
        execucao.iniciar_execucao()  # nova execução zera o pedido
        self.assertFalse(execucao.parada_pedida())

    def test_simulacao_grava_motivo_quando_linha_tentada_falha(self):
        planilha = PlanilhaVirtual()
        execucao.iniciar_execucao(simulacao=True)
        with linha_em_execucao("CE", planilha, 3, 9, 1, 1) as linha:
            linha.tentando()
            log_aviso("Linha 3: CPF 111: nenhum registro encontrado no SIGEF.")
        self.assertEqual(planilha.valor(3, 9), "SIMULADO ✗ CPF 111: nenhum registro encontrado no SIGEF.")

    def test_simulacao_nao_grava_em_linha_vazia(self):
        planilha = PlanilhaVirtual()
        execucao.iniciar_execucao(simulacao=True)
        with linha_em_execucao("CE", planilha, 3, 9, 1, 1):
            log_aviso("Linha 3: CPF ou valor vazio. Pulando.")  # nem chegou a tentar
        self.assertEqual(planilha.valor(3, 9), "")

    def test_simulacao_ok(self):
        planilha = PlanilhaVirtual()
        execucao.iniciar_execucao(simulacao=True)
        with linha_em_execucao("NL", planilha, 4, 10, 1, 1) as linha:
            linha.tentando()
            linha.simulado_ok("empenho 1")
        self.assertEqual(planilha.valor(4, 10), "SIMULADO ✓ NL conferida - nada foi gerado (empenho 1)")

    def test_modo_normal_nao_escreve_nada_extra(self):
        planilha = PlanilhaVirtual()
        execucao.iniciar_execucao(simulacao=False)
        with linha_em_execucao("CE", planilha, 3, 9, 1, 1) as linha:
            linha.tentando()
            log_aviso("Linha 3: qualquer problema")
        self.assertEqual(planilha.valor(3, 9), "")
