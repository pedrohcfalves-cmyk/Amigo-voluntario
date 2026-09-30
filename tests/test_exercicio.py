# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""Virada de ano: URLs do SIGEF e aba certa do navegador (item 1)."""
import unittest

from amigo.constantes import (
    REGEX_DOCUMENTO_OB, REGEX_DOCUMENTO_PP, URL_CE_SIGEF, ano_do_exercicio, url_do_exercicio,
)
from amigo.navegador import _aba_esta_na_tela
from amigo.utils import extrair_numero_documento, normalizar_numero_documento


class TesteExercicio(unittest.TestCase):
    def test_ano_vem_da_data(self):
        self.assertEqual(ano_do_exercicio({"data": "15012027"}), 2027)
        self.assertEqual(ano_do_exercicio({"data": "31/12/2026"}), 2026)

    def test_ano_manual_tem_prioridade(self):
        self.assertEqual(ano_do_exercicio({"data": "15012027", "ano_sigef": "2026"}), 2026)

    def test_sem_data_usa_ano_atual(self):
        from datetime import datetime
        self.assertEqual(ano_do_exercicio({}), datetime.now().year)
        self.assertEqual(ano_do_exercicio({"data": "99999999"}), datetime.now().year)

    def test_urls_trocam_de_ano(self):
        self.assertIn("/SIGEF2027/FIN/FINManterDespesaCertificada", url_do_exercicio(URL_CE_SIGEF, {"data": "02012027"}))
        self.assertNotIn("{ano}", url_do_exercicio(URL_CE_SIGEF, {}))
        antiga = "http://sigef.sefin.ro.gov.br/SIGEF2026/FIN/X.aspx"
        self.assertEqual(url_do_exercicio(antiga, {"data": "02012028"}), "http://sigef.sefin.ro.gov.br/SIGEF2028/FIN/X.aspx")

    def test_aba_de_outro_ano_nao_e_reaproveitada(self):
        desejada = "http://sigef/SIGEF2027/FIN/FINManterDespesaCertificada.aspx"
        self.assertTrue(_aba_esta_na_tela("http://sigef/SIGEF2027/FIN/FINManterDespesaCertificada.aspx?x=1", desejada, "FINManterDespesaCertificada"))
        self.assertFalse(_aba_esta_na_tela("http://sigef/SIGEF2026/FIN/FINManterDespesaCertificada.aspx", desejada, "FINManterDespesaCertificada"))
        self.assertFalse(_aba_esta_na_tela("http://sigef/SIGEF2027/FIN/Outra.aspx", desejada, "FINManterDespesaCertificada"))

    def test_documentos_de_2027_sao_reconhecidos(self):
        self.assertEqual(REGEX_DOCUMENTO_PP.search("O número gerado foi 2027PP000123.").group(0), "2027PP000123")
        self.assertEqual(REGEX_DOCUMENTO_OB.search("gerado 2027OB000001").group(0), "2027OB000001")
        self.assertEqual(extrair_numero_documento("2027CE018487"), "018487")
        self.assertEqual(normalizar_numero_documento("2027NL065995"), "65995")


if __name__ == "__main__":
    unittest.main()
