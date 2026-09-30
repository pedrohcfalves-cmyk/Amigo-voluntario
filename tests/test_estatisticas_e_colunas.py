# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""Estatísticas (tempo economizado, regressão) e Colunas Retráteis."""
from datetime import datetime, timedelta

from amigo import colunas, estatisticas, relatorio

from tests.apoio import TesteIsolado


class TesteEstatisticas(TesteIsolado):
    def test_tempo_economizado_e_regressao(self):
        base = datetime(2027, 1, 10, 8, 0, 0)
        for i, (itens, segundos) in enumerate([(5, 60), (10, 100), (20, 190)]):
            inicio = base + timedelta(hours=i)
            relatorio.registrar_execucao("CE", inicio, inicio + timedelta(seconds=segundos), itens, itens, True)
        dados = estatisticas.calcular_estatisticas("Todo o período", {})
        ce = dados["por_categoria"]["CE"]
        self.assertEqual(ce["processados"], 35)
        self.assertEqual(ce["duracao_manual_segundos"], 35 * 3 * 60)
        self.assertAlmostEqual(ce["regressao"][0], 15.0)
        self.assertAlmostEqual(ce["regressao"][1], 8.7142857, places=5)
        self.assertEqual(dados["totais"]["tempo_economizado_segundos"], 35 * 180 - 350)

    def test_filtro_por_mes(self):
        relatorio.registrar_execucao("NL", datetime(2026, 12, 1), datetime(2026, 12, 1, 0, 1), 3, 3, True)
        relatorio.registrar_execucao("NL", datetime(2027, 1, 5), datetime(2027, 1, 5, 0, 1), 4, 4, True)
        self.assertEqual(estatisticas.calcular_estatisticas("01/2027", {})["totais"]["processados"], 4)
        self.assertEqual(estatisticas.calcular_estatisticas("Todo o período", {})["totais"]["processados"], 7)


class TesteColunas(TesteIsolado):
    def test_padrao_de_fabrica(self):
        self.assertEqual(colunas.COL_CE_CPF, 1)   # B (base 0)
        self.assertEqual(colunas.COL_CE_GERADA, 9)  # I (base 1)

    def test_colisao_detectada(self):
        propostas = dict(colunas.COLUNAS_PADRAO, valor="B")
        self.assertIn("B", colunas._detectar_colisoes(propostas))
        self.assertEqual(colunas._detectar_colisoes(dict(colunas.COLUNAS_PADRAO)), {})

    def test_letras_e_indices(self):
        self.assertEqual(colunas.letra_para_indice_coluna("AA"), 27)
        self.assertEqual(colunas.indice_para_letra_coluna(27), "AA")
