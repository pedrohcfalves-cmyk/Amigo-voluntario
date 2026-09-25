"""
E/S pura de config.json: só o caminho do arquivo e a função de gravação
(`salvar_configuracoes`), sem depender de `CONFIG_PADRAO` nem de nenhuma
lógica de mesclagem - por isso é um módulo "folha", sem dependências
internas, que tanto `amigo.config` (parâmetros gerais) quanto
`amigo.colunas` (submenu de colunas) podem importar sem criar um ciclo
entre os dois.
"""
import json
import os

from .log import log_erro

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


def salvar_configuracoes(config: dict):
    """Salva o dicionário de configurações em config.json."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as arquivo:
            json.dump(config, arquivo, ensure_ascii=False, indent=4)
    except OSError as erro:
        log_erro(f"Falha ao salvar configurações: {erro}")
