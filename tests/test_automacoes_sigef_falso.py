# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
Roda o código REAL das automações CE, NL, PP e OB contra o "SIGEF de
mentira" (tests/sigef_falso.py) e confere:

  - modo normal: o clique final acontece e o número (2027CE..., 2027NL...)
    é gravado; a tela aberta é a do exercício certo (/SIGEF2027/);
  - modo simulação: o clique final NUNCA acontece e fica "SIMULADO ✓/✗";
  - botão Parar e progresso "linha X de Y";
  - textos de observação (padrão e personalizado).
"""
import importlib
from unittest import mock

from amigo import colunas, execucao

# importlib: `amigo.automacoes.ce` como atributo do pacote é a FUNÇÃO ce (o
# __init__ do pacote reexporta as funções com o mesmo nome dos módulos).
modulo_ce = importlib.import_module("amigo.automacoes.ce")
modulo_nl = importlib.import_module("amigo.automacoes.nl")
modulo_pp = importlib.import_module("amigo.automacoes.pp")
modulo_ob = importlib.import_module("amigo.automacoes.ob")
from amigo.lancamento_manual import PlanilhaVirtual, _montar_linha, executar_cadeia, montar_item

from tests.apoio import TesteIsolado
from tests.sigef_falso import SigefFalso, sync_playwright_falso

CONFIG = {"data": "15012027", "mes_referencia": "12/2026", "processo": "0029.000001/2027-10", "linha_inicial": 1}
DADOS = dict(cpf="529.982.247-25", ne="2027NE001234", valor="1.400,50", banco="1", agencia="1178-9", conta="74.508-1")


def novo_item(**documentos):
    item, erros = montar_item(DADOS)
    assert not erros, erros
    item.update(documentos)
    return item


class _Base(TesteIsolado):
    modulo = None
    funcao = None
    coluna = None

    def rodar(self, item, simulacao=False, sigef=None, config=None, linhas=1):
        sigef = sigef or SigefFalso()
        planilha = PlanilhaVirtual()
        dados = [_montar_linha(item) for _ in range(linhas)]
        with mock.patch.object(self.modulo, "sync_playwright", sync_playwright_falso(sigef)):
            execucao.iniciar_execucao(simulacao=simulacao)
            try:
                retorno = getattr(self.modulo, self.funcao)(dados, dict(config or CONFIG), planilha)
            finally:
                execucao.finalizar_execucao()
        return retorno, planilha.valor(1, getattr(colunas, self.coluna)), sigef


class TesteCE(_Base):
    modulo, funcao, coluna = modulo_ce, "ce", "COL_CE_GERADA"

    def test_modo_normal_gera_e_usa_exercicio_2027(self):
        retorno, valor, sigef = self.rodar(novo_item())
        self.assertEqual(retorno, 1)
        self.assertEqual(valor, "2027CE000111")
        self.assertIn("#btnManutencao_BtnIncluir", sigef.cliques())
        self.assertTrue(any("/SIGEF2027/" in url and "FINManterDespesaCertificada" in url for url in sigef.urls_abertas))

    def test_simulacao_nao_clica_em_incluir(self):
        retorno, valor, sigef = self.rodar(novo_item(), simulacao=True)
        self.assertNotIn("#btnManutencao_BtnIncluir", sigef.cliques())
        self.assertTrue(valor.startswith("SIMULADO ✓"), valor)
        self.assertEqual(retorno, 1)

    def test_simulacao_mostra_o_motivo_quando_daria_errado(self):
        _retorno, valor, sigef = self.rodar(novo_item(), simulacao=True, sigef=SigefFalso(credor_existe=False))
        self.assertTrue(valor.startswith("SIMULADO ✗"), valor)
        self.assertIn("529.982.247-25", valor)
        self.assertNotIn("#btnManutencao_BtnIncluir", sigef.cliques())

    def test_observacao_padrao_e_personalizada(self):
        _r, _v, sigef = self.rodar(novo_item())
        observacoes = [a[2] for a in sigef.acoes if a[0] == "fill" and a[1] == "#txtDeObservacao"]
        self.assertEqual(observacoes, ["Gratificação amigo voluntario 12/2026 Processo: 0029.000001/2027-10"])

        config = dict(CONFIG, observacao_ce="Pagamento {mes} - proc. {processo} {x}")
        _r, _v, sigef = self.rodar(novo_item(), config=config)
        observacoes = [a[2] for a in sigef.acoes if a[0] == "fill" and a[1] == "#txtDeObservacao"]
        self.assertEqual(observacoes, ["Pagamento 12/2026 - proc. 0029.000001/2027-10 {x}"])

    def test_botao_parar_antes_de_comecar(self):
        sigef = SigefFalso()
        with mock.patch.object(modulo_ce, "sync_playwright", sync_playwright_falso(sigef)):
            execucao.iniciar_execucao()
            execucao.pedir_parada()
            try:
                retorno = modulo_ce.ce([_montar_linha(novo_item())] * 3, dict(CONFIG), PlanilhaVirtual())
            finally:
                execucao.finalizar_execucao()
        self.assertEqual(retorno, 0)
        self.assertNotIn("#btnManutencao_BtnIncluir", sigef.cliques())

    def test_progresso_linha_x_de_y(self):
        recebidos = []
        ouvinte = lambda atual, total: recebidos.append((atual, total))
        execucao.adicionar_ouvinte_progresso(ouvinte)
        try:
            self.rodar(novo_item(), linhas=2)
        finally:
            execucao.remover_ouvinte_progresso(ouvinte)
        self.assertEqual(recebidos, [(1, 2), (2, 2)])


class TesteNL(_Base):
    modulo, funcao, coluna = modulo_nl, "nl", "COL_NL_GERADA"

    def test_modo_normal_gera(self):
        _r, valor, sigef = self.rodar(novo_item(ce="2027CE000111"))
        self.assertEqual(valor, "2027NL000222")
        self.assertIn("#btnRetencoesId", sigef.cliques())

    def test_simulacao_para_antes_das_retencoes_e_confirmacao(self):
        _r, valor, sigef = self.rodar(novo_item(ce="2027CE000111"), simulacao=True)
        self.assertNotIn("#btnRetencoesId", sigef.cliques())
        self.assertTrue(valor.startswith("SIMULADO ✓"), valor)

    def test_ce_simulada_nao_e_pesquisada(self):
        _r, valor, sigef = self.rodar(novo_item(ce="SIMULADO ✓ CE conferida - nada foi gerado"))
        self.assertEqual(valor, "")
        self.assertNotIn("#btnPesquisar", sigef.cliques())


class TestePP(_Base):
    modulo, funcao, coluna = modulo_pp, "pp", "COL_PP_GERADA"

    def test_modo_normal_gera_com_sigla_do_exercicio(self):
        _r, valor, sigef = self.rodar(novo_item(ce="000111", nl="000222"))
        self.assertEqual(valor, "2027PP000333")
        self.assertIn("#btnRetencoes", sigef.cliques())
        siglas = [a[2] for a in sigef.acoes if a[0] == "fill" and a[1] == "#txtNotaLancamentoSigla"]
        self.assertEqual(siglas, ["2027"])  # NL sem prefixo: usa o ano do exercício

    def test_simulacao_para_antes_das_retencoes(self):
        _r, valor, sigef = self.rodar(novo_item(ce="2027CE000111", nl="2027NL000222"), simulacao=True)
        self.assertNotIn("#btnRetencoes", sigef.cliques())
        self.assertTrue(valor.startswith("SIMULADO ✓"), valor)


class TesteOB(_Base):
    modulo, funcao, coluna = modulo_ob, "gerar_ob", "COL_OB_GERADA"

    def test_modo_normal_gera(self):
        retorno, valor, sigef = self.rodar(novo_item(pp="2027PP000333"))
        self.assertEqual(valor, "2027OB000444")
        self.assertEqual(retorno, (1, 1))
        self.assertIn("#SIGEFBotoesManutencao_BtnIncluir", sigef.cliques())
        observacoes = [a[2] for a in sigef.acoes if a[0] == "fill" and a[1] == "#txtDeObservacao"]
        self.assertEqual(observacoes, ["Ressarcimento de amigos voluntario Referente ao 12/2026 Processo: 0029.000001/2027-10"])

    def test_simulacao_nao_confirma_lote_nem_inclui(self):
        _r, valor, sigef = self.rodar(novo_item(pp="2027PP000333"), simulacao=True)
        self.assertNotIn("#btnConfirmar", sigef.cliques())
        self.assertNotIn("#SIGEFBotoesManutencao_BtnIncluir", sigef.cliques())
        self.assertTrue(valor.startswith("SIMULADO ✓"), valor)


class TesteCadeiaSemPlanilhaComSigefFalso(TesteIsolado):
    """O modo Sem planilha rodando as 4 automações reais em sequência."""

    def _rodar_cadeia(self, simulacao):
        sigef = SigefFalso()
        fabrica = sync_playwright_falso(sigef)
        item = novo_item()
        with mock.patch.object(modulo_ce, "sync_playwright", fabrica), \
                mock.patch.object(modulo_nl, "sync_playwright", fabrica), \
                mock.patch.object(modulo_pp, "sync_playwright", fabrica), \
                mock.patch.object(modulo_ob, "sync_playwright", fabrica):
            execucao.iniciar_execucao(simulacao=simulacao)
            try:
                executar_cadeia(item, 1, dict(CONFIG))
            finally:
                execucao.finalizar_execucao()
        return item, sigef

    def test_cadeia_de_verdade_gera_os_4_numeros_de_2027(self):
        item, _sigef = self._rodar_cadeia(simulacao=False)
        self.assertEqual(item["situacao"], "concluido", item.get("motivo"))
        self.assertEqual(
            (item["ce"], item["nl"], item["pp"], item["ob"]),
            ("2027CE000111", "2027NL000222", "2027PP000333", "2027OB000444"),
        )

    def test_cadeia_simulada_para_na_ce_sem_gerar_nada(self):
        item, sigef = self._rodar_cadeia(simulacao=True)
        self.assertEqual(item["situacao"], "simulado")
        self.assertTrue(item["ce"].startswith("SIMULADO ✓"))
        self.assertEqual((item["nl"], item["pp"], item["ob"]), ("", "", ""))
        self.assertNotIn("#btnManutencao_BtnIncluir", sigef.cliques())
