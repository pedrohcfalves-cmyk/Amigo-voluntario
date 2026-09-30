#!/usr/bin/env python3
# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
==============================================================================
SISTEMA DE AUTOMAÇÕES SIGEF - INTERFACE GRÁFICA
==============================================================================
Ponto de entrada da versão com janela (para quem não quer usar o terminal).
Toda a lógica vive no pacote `amigo/` (veja o README.md); este arquivo só
chama `amigo.gui.iniciar()`.

Ao abrir, o aplicativo verifica se falta algo para funcionar (pacotes
Python, navegador do Playwright) e instala automaticamente antes de
liberar as automações.

Execução:
    python main_gui.py
ou, no Windows, dando dois cliques em Amigo.bat.

O menu de terminal continua disponível em `main.py` / `python main.py`,
sem nenhuma mudança de comportamento.
==============================================================================
"""
from amigo.gui import iniciar

if __name__ == "__main__":
    iniciar()
