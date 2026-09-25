"""
MENUS (principal e de programas) + PONTO DE ENTRADA
=====================================================
Junta as peças dos outros módulos numa interface de terminal: menu
principal (parâmetros, planilha, colunas, executar automação) e o
submenu de automações (dicionário `amigo.automacoes.PROGRAMAS`).
"""
import sys
import traceback

from .automacoes import PROGRAMAS
from .colunas import menu_configurar_colunas
from .config import (
    carregar_configuracoes, configurar_parametros, alterar_planilha,
    alterar_linha_inicial, exibir_configuracoes,
)
from .excel import conectar_planilha, ler_dados
from .log import log_info, log_sucesso, log_erro, log_aviso


def print_menu():
    """Exibe o menu principal do sistema."""
    print("\n🚀 SISTEMA DE AUTOMAÇÕES SIGEF 🚀")
    print("=" * 50)
    print("1 - Configurar parâmetros")
    print("2 - Executar automação")
    print("3 - Alterar planilha")
    print("4 - Alterar linha inicial")
    print("5 - Configurar colunas (Colunas Retráteis)")
    print("6 - Exibir configurações atuais")
    print("7 - Sair")
    print("=" * 50)


def menu_programas() -> int:
    """Exibe o submenu com as automações disponíveis e retorna a opção escolhida."""
    print("\n🧩 MENU DE PROGRAMAS")
    print("=" * 50)
    for numero, (nome, _funcao) in PROGRAMAS.items():
        print(f"{numero:>2} - {nome}")
    print(" 0 - Voltar")
    print("=" * 50)

    escolha = input("Escolha uma automação: ").strip()
    if not escolha.isdigit():
        log_aviso("Opção inválida.")
        return 0

    return int(escolha)


def executar_automacao(config: dict):
    """
    Fluxo completo de execução de uma automação:
        1. Conecta ao arquivo Excel via COM (abre/reaproveita, de forma
           visível, para acompanhamento em tempo real)
        2. Lê os dados a partir da linha inicial configurada
        3. Exibe o submenu de programas
        4. Executa a função correspondente via dicionário PROGRAMAS
    """
    if not config.get("caminho_planilha"):
        log_erro("Nenhum arquivo Excel configurado. Use a opção 'Alterar planilha' primeiro.")
        return

    log_info("Conectando ao arquivo Excel...")
    worksheet = conectar_planilha(config)
    if worksheet is None:
        log_erro("Não foi possível continuar sem conexão com a planilha.")
        return

    linha_inicial = config.get("linha_inicial", 2)
    dados = ler_dados(worksheet, linha_inicial)

    if not dados:
        log_aviso("Nenhum dado encontrado para processar.")
        return

    opcao = menu_programas()
    if opcao == 0:
        log_info("Retornando ao menu principal.")
        return

    if opcao not in PROGRAMAS:
        log_erro("Opção de automação inválida.")
        return

    nome_programa, funcao = PROGRAMAS[opcao]
    log_info(f"Iniciando automação: {nome_programa}...")

    try:
        funcao(dados, config, worksheet)
        log_sucesso(f"Automação '{nome_programa}' finalizada.")
    except Exception:
        log_erro(f"Erro inesperado durante a automação '{nome_programa}':")
        traceback.print_exc()


def menu_principal():
    """Loop principal do sistema, controla a navegação entre as opções."""
    config = carregar_configuracoes()

    while True:
        print_menu()
        opcao = input("Escolha uma opção: ").strip()

        try:
            if opcao == "1":
                config = configurar_parametros(config)

            elif opcao == "2":
                executar_automacao(config)

            elif opcao == "3":
                config = alterar_planilha(config)

            elif opcao == "4":
                config = alterar_linha_inicial(config)

            elif opcao == "5":
                config = menu_configurar_colunas(config)

            elif opcao == "6":
                exibir_configuracoes(config)

            elif opcao == "7":
                log_info("Encerrando o sistema. Até logo! 👋")
                sys.exit(0)

            else:
                log_aviso("Opção inválida. Tente novamente.")

        except KeyboardInterrupt:
            print()
            log_info("Execução interrompida pelo usuário. Encerrando...")
            sys.exit(0)
        except Exception:
            log_erro("Ocorreu um erro inesperado:")
            traceback.print_exc()


def main():
    """Ponto de entrada da aplicação."""
    try:
        menu_principal()
    except Exception:
        log_erro("Erro fatal na aplicação:")
        traceback.print_exc()
        sys.exit(1)
