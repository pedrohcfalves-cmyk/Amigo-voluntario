"""
Carregar/salvar config.json e os menus interativos de parâmetros gerais
(data, processo, mês de referência, linha inicial, arquivo Excel). A
configuração das colunas em si (submenu, tabela, validação) vive em
`amigo.colunas` - este módulo só cuida de mesclar a chave "colunas" com o
padrão de fábrica ao carregar, e de aplicá-la via `colunas.aplicar_colunas()`.
"""
import json
import os

from . import colunas
from .colunas import aplicar_colunas, exibir_colunas_atuais
from .config_io import BASE_DIR, CONFIG_PATH, salvar_configuracoes
from .log import log_sucesso, log_erro, log_aviso

# Configuração padrão utilizada caso ainda não exista config.json
CONFIG_PADRAO = {
    "data": "",                 # Ex: 30062026
    "processo": "",             # Ex: 0029.037004/2026-72
    "mes_referencia": "",       # Ex: 06/2026 (sempre o mês anterior!)
    "linha_inicial": 2,         # Ex: 37
    "caminho_planilha": "",     # Ex: C:\\Pasta\\SIGEF\\controle.xlsx
    "aba_planilha": "",         # Nome da aba (opcional; se vazio, usa a 1ª aba)

    # "Colunas Retráteis": em qual coluna da planilha do usuário está (ou
    # deve ser salvo) cada dado. O padrão de fábrica vem de
    # `colunas.COLUNAS_PADRAO` (fonte única, ver amigo/colunas.py).
    "colunas": dict(colunas.COLUNAS_PADRAO),
}


def carregar_configuracoes() -> dict:
    """
    Carrega as configurações do arquivo config.json.
    Caso o arquivo não exista, cria um com valores padrão.

    Sempre que uma configuração é carregada, `aplicar_colunas()` é chamada
    para que as constantes COL_* (usadas por todas as automações) fiquem de
    acordo com as colunas configuradas pelo usuário - mesmo que config.json
    seja de uma versão anterior do sistema (sem a chave "colunas") ou tenha
    sido editado manualmente.
    """
    if not os.path.exists(CONFIG_PATH):
        salvar_configuracoes(CONFIG_PADRAO)
        config = json.loads(json.dumps(CONFIG_PADRAO))  # cópia independente (inclui "colunas")
        aplicar_colunas(config["colunas"])
        return config

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as arquivo:
            config = json.load(arquivo)
            # Garante que todas as chaves padrão existam (versões antigas do arquivo)
            for chave, valor in CONFIG_PADRAO.items():
                if chave == "colunas":
                    continue
                config.setdefault(chave, valor)

            # "colunas" é mesclada campo a campo (não com setdefault direto)
            # para que um config.json antigo - sem essa chave, ou com só
            # ALGUMAS colunas configuradas - ganhe o padrão de fábrica
            # automaticamente nas que faltarem, sem perder as que o usuário
            # já tiver configurado.
            colunas = config.get("colunas")
            if not isinstance(colunas, dict):
                colunas = {}
            for chave, valor in CONFIG_PADRAO["colunas"].items():
                colunas.setdefault(chave, valor)
            config["colunas"] = colunas

            aplicar_colunas(colunas)
            return config
    except (json.JSONDecodeError, OSError) as erro:
        log_erro(f"Não foi possível ler config.json ({erro}). Recriando com valores padrão.")
        salvar_configuracoes(CONFIG_PADRAO)
        config = json.loads(json.dumps(CONFIG_PADRAO))
        aplicar_colunas(config["colunas"])
        return config


# ==============================================================================
# MENUS INTERATIVOS DE CONFIGURAÇÃO
# ==============================================================================

def configurar_parametros(config: dict) -> dict:
    """
    Menu interativo para o usuário informar/alterar:
    DATA, PROCESSO, MÊS REFERÊNCIA e LINHA INICIAL.
    O arquivo Excel (caminho/aba) é configurado separadamente em
    'alterar_planilha'.
    """
    print("\n🔧 CONFIGURAR PARÂMETROS")
    print("-" * 50)

    data = input(f"DATA (ex: 30062026) [{config.get('data') or 'vazio'}]: ").strip()
    if data:
        config["data"] = data

    processo = input(
        f"PROCESSO (ex: 0029.037004/2026-72) [{config.get('processo') or 'vazio'}]: "
    ).strip()
    if processo:
        config["processo"] = processo

    mes_referencia = input(
        f"MÊS REFERÊNCIA (ex: 06/2026 | Sempre o mês anterior!) [{config.get('mes_referencia') or 'vazio'}]: "
    ).strip()
    if mes_referencia:
        config["mes_referencia"] = mes_referencia

    linha_inicial = input(
        f"LINHA INICIAL DA LEITURA (ex: 37) [{config.get('linha_inicial')}]: "
    ).strip()
    if linha_inicial:
        if linha_inicial.isdigit():
            config["linha_inicial"] = int(linha_inicial)
        else:
            log_aviso("Linha inicial inválida. Mantendo valor anterior.")

    salvar_configuracoes(config)
    log_sucesso("Parâmetros atualizados com sucesso!")
    return config


def alterar_planilha(config: dict) -> dict:
    """Menu interativo para alterar o caminho do arquivo Excel e o nome da aba."""
    print("\n📊 ALTERAR ARQUIVO EXCEL")
    print("-" * 50)

    caminho_planilha = input(
        f"Caminho do arquivo .xlsx [{config.get('caminho_planilha') or 'vazio'}]: "
    ).strip().strip('"')
    if caminho_planilha:
        config["caminho_planilha"] = caminho_planilha

    aba_planilha = input(
        f"Nome da aba (opcional, deixe vazio para usar a 1ª aba) [{config.get('aba_planilha') or 'vazio'}]: "
    ).strip()
    if aba_planilha:
        config["aba_planilha"] = aba_planilha

    salvar_configuracoes(config)
    log_sucesso("Arquivo Excel atualizado com sucesso!")
    return config


def alterar_linha_inicial(config: dict) -> dict:
    """Menu interativo dedicado apenas à troca rápida da linha inicial."""
    print("\n🔢 ALTERAR LINHA INICIAL")
    print("-" * 50)

    linha_inicial = input(f"Nova linha inicial [{config.get('linha_inicial')}]: ").strip()
    if linha_inicial.isdigit():
        config["linha_inicial"] = int(linha_inicial)
        salvar_configuracoes(config)
        log_sucesso(f"Linha inicial alterada para {linha_inicial}.")
    else:
        log_aviso("Valor inválido. Nenhuma alteração foi feita.")

    return config


def exibir_configuracoes(config: dict):
    """Exibe as configurações atuais de forma organizada."""
    print("\n📋 CONFIGURAÇÕES ATUAIS")
    print("-" * 50)
    print(f"DATA .................: {config.get('data') or '(não definido)'}")
    print(f"PROCESSO .............: {config.get('processo') or '(não definido)'}")
    print(f"MÊS REFERÊNCIA .......: {config.get('mes_referencia') or '(não definido)'}")
    print(f"LINHA INICIAL ........: {config.get('linha_inicial')}")
    print(f"ARQUIVO EXCEL ........: {config.get('caminho_planilha') or '(não definido)'}")
    print(f"ABA ..................: {config.get('aba_planilha') or '(1ª aba)'}")
    print("-" * 50)
    exibir_colunas_atuais(config)

