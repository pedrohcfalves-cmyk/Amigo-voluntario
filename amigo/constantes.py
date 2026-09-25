"""
Constantes do sistema: ambiente SIGEF (homologação/produção), URLs de cada
tela, timeouts e regex pré-compiladas. NADA aqui é configurável pelo
usuário em tempo de execução (isso é config.json / colunas.py) - são
constantes do PROGRAMA, versionadas junto com o código.
"""
import re

DOMINIO_SIGEF = "sigef.sefin.ro.gov.br"  # <-- ambiente de PRODUÇÃO

# ==============================================================================
# URLs DO SIGEF
# ==============================================================================
# Como as demais automações (NL, PP, OB) ainda serão programadas, deixe
# aqui as URLs de cada tela do SIGEF assim que você definir cada fluxo -
# por exemplo:
#
#   URL_NL_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/....aspx"
#   URL_PP_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/....aspx"
#   URL_OB_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/....aspx"

# Tela "Manter Despesa Certificada" (CE).
URL_CE_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/FINManterDespesaCertificada.aspx?CdTransacao=121"
MARCADOR_URL_CE = "FINManterDespesaCertificada"

# Tela "Listar Despesa Certificada Geral" (Raspar CE).
URL_RASPAR_CE_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/FINListarDespesaCertificadaGeral.aspx?CdTransacao=123"
MARCADOR_URL_RASPAR_CE = "FINListarDespesaCertificadaGeral"

# Tela "Liquidar Despesa Certificada" (NL).
URL_NL_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/FINLiquidarDespesaCertificada.aspx?CdTransacao=160"
MARCADOR_URL_NL = "FINLiquidarDespesaCertificada"

# Tela "Preparação Pagamento Despesa Empenhada" (PP).
URL_PP_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/FINPreparacaoPagamentoDespesaEmpenhada.aspx?CdTransacao=250"
MARCADOR_URL_PP = "FINPreparacaoPagamentoDespesaEmpenhada"

# Tela "Listar Preparação Pagamento Geral" (Raspar contas).
URL_RASPAR_CONTA_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/FINListarPreparacaoPagamentoGeral.aspx?CdTransacao=181"
MARCADOR_URL_RASPAR_CONTA = "FINListarPreparacaoPagamentoGeral"

# Tela "Manter Ordem Bancária" (Gerar OB, lote de 30 em 30, fluxo
# Descentralizada / Tipo de OB "2").
URL_OB_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF2026/FIN/FINManterOrdemBancaria.aspx?CdTransacao=214"
MARCADOR_URL_OB = "FINManterOrdemBancaria"

TIMEOUT_PADRAO_SIGEF = 30000  # 30s

# Espera curta para coisas que "podem ou não" acontecer - hoje só o popup do
# SIGEF que se fecha sozinho após a seleção do item. Curto de propósito: se o
# popup não fechar, `fechar_paginas()` termina o serviço sem custar 30s.
TIMEOUT_POPUP_FECHAR = 800  # 0,8s

# Espera (e repetições) do clique de seleção em grade de popup: o SIGEF às
# vezes engole o 1º clique e a tela principal continua sem o registro - o que
# depois fazia a comparação do domicílio bancário rodar com os dados errados.
TIMEOUT_CLIQUE_GRADE = 8000  # 8s por tentativa

# Espera pelo botão "Limpar". Curta de propósito: quando ele não está na tela,
# o certo é voltar (`#btnVoltar`) e tentar de novo, não ficar 30s parado.
TIMEOUT_BOTAO_LIMPAR = 5000  # 5s
TENTATIVAS_CLIQUE_GRADE = 3

# Tamanho do lote da automação "Gerar OB" (quantas linhas da planilha por
# OB) e trava de segurança contra loop infinito de paginação dentro de 1
# lote (ver `_montar_lote_ob`).
TAMANHO_LOTE_OB = 30
MAX_PAGINAS_LOTE_OB = 40

# Tipo de Ordem Bancária usado pela PP. Selecionado por VALUE (e não por
# índice): o <select> tem uma 1ª opção em branco, então qualquer mudança de
# ordem/opção no SIGEF alteraria silenciosamente a escolha feita por índice.
#   ""=(em branco)  1=Centralizada  2=Descentralizada  3=Regularização
TIPO_ORDEM_BANCARIA_PP = "2"  # Descentralizada

# ---- Regex pré-compiladas (compiladas 1 única vez, no import do módulo) --
REGEX_APENAS_DIGITOS = re.compile(r"\D")
REGEX_PREFIXO_DOCUMENTO = re.compile(r"^\d{4}(PP|OB|CE|NL|NE)")
# Usada pela PP para extrair o número gerado da mensagem de sucesso do SIGEF
# ("... O número gerado foi 2026PP039772.").
REGEX_DOCUMENTO_PP = re.compile(r"\d{4}PP\d+", re.IGNORECASE)
# Usada pela Gerar OB para extrair o número gerado da mensagem de sucesso
# do SIGEF ("... O número gerado foi 2026OB012345.").
REGEX_DOCUMENTO_OB = re.compile(r"\d{4}OB\d+", re.IGNORECASE)
# Todos os números soltos no texto de uma linha de grade (ex: de
# "2026NL065995" saem "2026" e "065995") - usada para identificar qual linha
# corresponde aos documentos da planilha.
REGEX_NUMEROS = re.compile(r"\d+")
# Célula de grade que É um valor monetário no padrão BR ("840,00",
# "1.260,50") - usada pela Raspar PP para não tentar converter em número
# células como "2026PP053126" ou o nome do favorecido.
REGEX_VALOR_BR = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$")
