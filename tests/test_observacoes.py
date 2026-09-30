# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""Textos de observação da CE e da OB (item 6)."""
import unittest

from amigo.observacoes import TEXTO_PADRAO_OBSERVACAO, aplicar_marcacoes, modelo_observacao, montar_observacao

CONFIG = {"mes_referencia": "12/2026", "processo": "0029.1/2027-10"}


class TesteObservacoes(unittest.TestCase):
    def test_texto_padrao_continua_igual_ao_de_sempre(self):
        self.assertEqual(montar_observacao("ce", CONFIG), "Gratificação amigo voluntario 12/2026 Processo: 0029.1/2027-10")
        self.assertEqual(montar_observacao("ob", CONFIG),
                         "Ressarcimento de amigos voluntario Referente ao 12/2026 Processo: 0029.1/2027-10")

    def test_texto_personalizado(self):
        config = dict(CONFIG, observacao_ce="Pagamento de {mes} ({processo})")
        self.assertEqual(montar_observacao("ce", config), "Pagamento de 12/2026 (0029.1/2027-10)")
        self.assertEqual(montar_observacao("ob", config), montar_observacao("ob", CONFIG))  # OB continua padrão

    def test_vazio_ou_espacos_volta_ao_padrao(self):
        self.assertEqual(modelo_observacao("ce", {"observacao_ce": "   "}), TEXTO_PADRAO_OBSERVACAO["ce"])

    def test_chaves_desconhecidas_nao_quebram(self):
        self.assertEqual(aplicar_marcacoes("{mes} {outra} {}", "01/2027", "p"), "01/2027 {outra} {}")


if __name__ == "__main__":
    unittest.main()
