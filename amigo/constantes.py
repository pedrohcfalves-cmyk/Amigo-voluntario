# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
Constantes do sistema: ambiente SIGEF (homologação/produção), URLs de cada
tela, timeouts e regex pré-compiladas. NADA aqui é configurável pelo
usuário em tempo de execução (isso é config.json / colunas.py) - são
constantes do PROGRAMA, versionadas junto com o código.
"""
import re

DOMINIO_SIGEF = "sigef.sefin.ro.gov.br"  # <-- ambiente de PRODUÇÃO

# ==============================================================================
# URLs DO SIGEF (com o ANO DO EXERCÍCIO)
# ==============================================================================
# O endereço de cada tela do SIGEF muda todo ano: /SIGEF2026/, /SIGEF2027/...
# Por isso as URLs abaixo são MODELOS com "{ano}" no lugar do ano - quem abre
# a tela usa `url_do_exercicio(URL_..., config)`, que troca "{ano}" pelo ano
# certo (ver `ano_do_exercicio`). Assim, na virada de 2026 para 2027, nada
# precisa ser alterado no código: os documentos passam a sair como
# 2027CE..., 2027NL..., 2027PP..., 2027OB... (as regex mais abaixo já
# aceitam qualquer ano).

# Tela "Manter Despesa Certificada" (CE).
URL_CE_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF{{ano}}/FIN/FINManterDespesaCertificada.aspx?CdTransacao=121"
MARCADOR_URL_CE = "FINManterDespesaCertificada"

# Tela "Listar Despesa Certificada Geral" (Raspar CE).
URL_RASPAR_CE_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF{{ano}}/FIN/FINListarDespesaCertificadaGeral.aspx?CdTransacao=123"
MARCADOR_URL_RASPAR_CE = "FINListarDespesaCertificadaGeral"

# Tela "Liquidar Despesa Certificada" (NL).
URL_NL_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF{{ano}}/FIN/FINLiquidarDespesaCertificada.aspx?CdTransacao=160"
MARCADOR_URL_NL = "FINLiquidarDespesaCertificada"

# Tela "Preparação Pagamento Despesa Empenhada" (PP).
URL_PP_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF{{ano}}/FIN/FINPreparacaoPagamentoDespesaEmpenhada.aspx?CdTransacao=250"
MARCADOR_URL_PP = "FINPreparacaoPagamentoDespesaEmpenhada"

# Tela "Listar Preparação Pagamento Geral" (Raspar contas).
URL_RASPAR_CONTA_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF{{ano}}/FIN/FINListarPreparacaoPagamentoGeral.aspx?CdTransacao=181"
MARCADOR_URL_RASPAR_CONTA = "FINListarPreparacaoPagamentoGeral"

# Tela "Manter Ordem Bancária" (Gerar OB, lote de 30 em 30, fluxo
# Descentralizada / Tipo de OB "2").
URL_OB_SIGEF = f"http://{DOMINIO_SIGEF}/SIGEF{{ano}}/FIN/FINManterOrdemBancaria.aspx?CdTransacao=214"
MARCADOR_URL_OB = "FINManterOrdemBancaria"

# Pedaço da URL que identifica o exercício (ex: "/SIGEF2027/").
REGEX_SEGMENTO_EXERCICIO = re.compile(r"/SIGEF(\d{4})/", re.IGNORECASE)


def ano_do_exercicio(config=None) -> int:
    """
    Ano do exercício do SIGEF usado nas URLs, nesta ordem de prioridade:

      1. `config["ano_sigef"]` - o campo "Ano do SIGEF" dos Parâmetros,
         quando preenchido (só para casos especiais, ex: lançar em janeiro
         ainda no exercício anterior);
      2. o ano da "Data" dos parâmetros (DDMMAAAA) - a data da operação é
         que define o exercício: uma Data 15012027 vai para o SIGEF2027;
      3. o ano atual do computador.
    """
    from datetime import datetime

    config = config or {}
    ano_manual = str(config.get("ano_sigef") or "").strip()
    if ano_manual.isdigit() and len(ano_manual) == 4:
        return int(ano_manual)

    digitos = re.sub(r"\D", "", str(config.get("data") or ""))
    if len(digitos) == 8:
        try:
            return datetime(int(digitos[4:]), int(digitos[2:4]), int(digitos[:2])).year
        except ValueError:
            pass
    return datetime.now().year


def url_do_exercicio(modelo: str, config=None) -> str:
    """Troca "{ano}" do modelo de URL pelo ano do exercício (ver
    `ano_do_exercicio`). Também aceita uma URL já com ano fixo
    (/SIGEF2026/): o ano dela é substituído do mesmo jeito."""
    ano = ano_do_exercicio(config)
    if "{ano}" in modelo:
        return modelo.replace("{ano}", str(ano))
    return REGEX_SEGMENTO_EXERCICIO.sub(f"/SIGEF{ano}/", modelo)


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
