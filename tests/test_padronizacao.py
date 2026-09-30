# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""Padronização dos dados digitados no modo Sem planilha e nos Parâmetros."""
import unittest

from amigo import lancamento_manual as lm


class TestePadronizacao(unittest.TestCase):
    def ok(self, campo, texto, esperado):
        self.assertEqual(lm.padronizar_campo(campo, texto)[0], esperado, f"{campo}: {texto!r}")

    def erro(self, campo, texto):
        with self.assertRaises(ValueError, msg=f"{campo}: {texto!r}"):
            lm.padronizar_campo(campo, texto)

    def test_cpf(self):
        self.ok("cpf", "529.982.247-25", "52998224725")
        self.ok("cpf", " 529 982 247 25 ", "52998224725")
        for invalido in ("52998224724", "1234", "111.111.111-11", ""):
            self.erro("cpf", invalido)

    def test_nota_de_empenho(self):
        self.ok("ne", "2027NE1234", "2027NE001234")
        self.ok("ne", "2027 ne 1234", "2027NE001234")
        self.ok("ne", "1234", "001234")
        for invalido in ("20271234", "2027NE0", "12a", ""):
            self.erro("ne", invalido)

    def test_banco_agencia_conta(self):
        self.ok("banco", "001", "1")
        self.ok("banco", "0237", "237")
        self.erro("banco", "1234")
        self.erro("banco", "BB")
        self.ok("agencia", "2757x", "2757X")
        self.ok("agencia", "1.178-9", "1178-9")
        self.erro("agencia", "11-78-9")
        self.ok("conta", "74.508-1", "74.508-1")
        self.ok("conta", "81.590-Х", "81.590-X")  # X cirílico (copiado de PDF)
        self.erro("conta", "0000")

    def test_valor(self):
        casos = {"1400": "1400,00", "1400,50": "1400,50", "1.400,50": "1400,50",
                 "R$ 1.400": "1400,00", "1400.5": "1400,50", "0,01": "0,01"}
        for texto, esperado in casos.items():
            self.ok("valor", texto, esperado)
        for invalido in ("0", "1.40,00", "abc", "12,345", ""):
            self.erro("valor", invalido)

    def test_parametros(self):
        self.assertEqual(lm.padronizar_data("15/01/2027")[0], "15012027")
        self.assertEqual(lm.padronizar_mes_referencia("12/2026")[0], "12/2026")
        self.assertEqual(lm.padronizar_mes_referencia("6/2027")[0], "06/2027")
        self.assertEqual(lm.padronizar_processo(" 0029.037004 /2026-72 ")[0], "0029.037004/2026-72")
        for funcao, invalido in ((lm.padronizar_data, "31022027"), (lm.padronizar_mes_referencia, "13/2027"),
                                 (lm.padronizar_processo, "abc")):
            with self.assertRaises(ValueError):
                funcao(invalido)

    def test_aviso_mes_anterior_na_virada_do_ano(self):
        self.assertEqual(lm.aviso_mes_referencia("15012027", "12/2026"), "")
        self.assertIn("12/2026", lm.aviso_mes_referencia("15012027", "01/2027"))

    def test_montar_item_e_parametros(self):
        item, erros = lm.montar_item(dict(cpf="52998224725", ne="1", banco="1", agencia="1", conta="1", valor="10"))
        self.assertEqual(erros, {})
        self.assertEqual(item["exibicao"]["valor"], "R$ 10,00")
        parametros, erros = lm.montar_parametros_manuais(
            {"data": "15012027", "processo": "1/2027", "mes_referencia": "12/2026", "observacao_ce": "  Meu texto  "}
        )
        self.assertEqual(erros, {})
        self.assertEqual(parametros["observacao_ce"], "Meu texto")
        self.assertEqual(parametros["observacao_ob"], "")


if __name__ == "__main__":
    unittest.main()
