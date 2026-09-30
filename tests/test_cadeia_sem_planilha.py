# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""Cadeia CE -> NL -> PP -> OB do modo Sem planilha, com automações de
mentira (confere o repasse dos números entre as etapas)."""
from amigo import colunas
from amigo import lancamento_manual as lm
from amigo.excel import atualizar_status, obter_celula, salvar_valor_gerado

from tests.apoio import TesteIsolado

DADOS = dict(cpf="529.982.247-25", ne="2027NE1234", banco="001", agencia="1178-9", conta="74.508-1", valor="1.400,50")


def automacoes_falsas(chamadas, pp_falha=False):
    def ce(dados, config, ws):
        n = config["linha_inicial"]
        assert obter_celula(dados[0], colunas.COL_CE_CPF) == "52998224725"
        chamadas.append("CE")
        salvar_valor_gerado(ws, n, colunas.COL_CE_GERADA, "2027CE000111")
        return 1

    def nl(dados, config, ws):
        n = config["linha_inicial"]
        assert obter_celula(dados[0], colunas.COL_NL_CE) == "2027CE000111"
        chamadas.append("NL")
        salvar_valor_gerado(ws, n, colunas.COL_NL_GERADA, "2027NL000222")
        return 1

    def pp(dados, config, ws):
        n = config["linha_inicial"]
        assert obter_celula(dados[0], colunas.COL_PP_NL) == "2027NL000222"
        chamadas.append("PP")
        if pp_falha:
            atualizar_status(ws, n, colunas.COL_PP_GERADA, "Saldo insuficiente.")
            return 0
        salvar_valor_gerado(ws, n, colunas.COL_PP_GERADA, "2027PP000333")
        return 1

    def ob(dados, config, ws):
        n = config["linha_inicial"]
        assert obter_celula(dados[0], colunas.COL_OB_PP_ESPERADO) == "2027PP000333"
        chamadas.append("OB")
        salvar_valor_gerado(ws, n, colunas.COL_OB_GERADA, "2027OB000444")
        return (1, 1)

    return {"CE": ce, "NL": nl, "PP": pp, "OB": ob}


class TesteCadeia(TesteIsolado):
    def novo(self):
        item, erros = lm.montar_item(DADOS)
        self.assertEqual(erros, {})
        return item

    def test_caminho_feliz(self):
        chamadas = []
        item = lm.executar_cadeia(self.novo(), 7, {}, funcoes=automacoes_falsas(chamadas))
        self.assertEqual(item["situacao"], "concluido")
        self.assertEqual(chamadas, ["CE", "NL", "PP", "OB"])
        self.assertEqual(item["ob"], "2027OB000444")

    def test_para_na_pp_e_retoma_sem_refazer_ce_e_nl(self):
        chamadas = []
        item = lm.executar_cadeia(self.novo(), 1, {}, funcoes=automacoes_falsas(chamadas, pp_falha=True))
        self.assertEqual((item["situacao"], item["etapa_parada"], item["motivo"]), ("parou", "PP", "Saldo insuficiente."))
        chamadas.clear()
        item = lm.executar_cadeia(item, 1, {}, funcoes=automacoes_falsas(chamadas))
        self.assertEqual(item["situacao"], "concluido")
        self.assertEqual(chamadas, ["PP", "OB"])

    def test_colunas_em_outras_letras(self):
        colunas.aplicar_colunas({"cpf": "A", "nota_empenho": "C", "banco": "P", "agencia": "Q", "conta": "R",
                                 "valor": "S", "ce": "T", "nl": "U", "pp": "V", "ob": "W",
                                 "raspar_pp_observacao": "X", "raspar_conta_resultado": "X",
                                 "raspar_conta_conferencia": "Y"})
        item = lm.executar_cadeia(self.novo(), 1, {}, funcoes=automacoes_falsas([]))
        self.assertEqual(item["situacao"], "concluido")

    def test_nada_e_salvo_no_disco(self):
        import os
        from amigo import config_io, relatorio
        lm.executar_cadeia(self.novo(), 1, {}, funcoes=automacoes_falsas([]))
        self.assertFalse(os.path.exists(config_io.CONFIG_PATH))
        self.assertFalse(os.path.exists(relatorio.HISTORICO_PATH))
