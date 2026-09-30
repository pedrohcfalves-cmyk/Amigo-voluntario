# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
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
import sys

from .log import log_erro

if getattr(sys, "frozen", False):
    # Empacotado com PyInstaller (--onefile): `__file__` apontaria para a
    # pasta temporária de extração (_MEIPASS), que é apagada ao fechar o
    # programa - o config.json tem que ficar ao lado do .exe de verdade,
    # senão a configuração "some" a cada reabertura.
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")


def salvar_configuracoes(config: dict):
    """Salva o dicionário de configurações em config.json."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as arquivo:
            json.dump(config, arquivo, ensure_ascii=False, indent=4)
    except OSError as erro:
        log_erro(f"Falha ao salvar configurações: {erro}")
