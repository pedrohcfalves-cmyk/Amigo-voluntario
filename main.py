#!/usr/bin/env python3
# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
==============================================================================
SISTEMA DE AUTOMAÇÕES SIGEF
==============================================================================
Ponto de entrada. Toda a lógica vive no pacote `amigo/` (veja o README.md
para o mapa completo dos módulos); este arquivo só chama `amigo.menu.main()`.

Execução:
    python main.py

Requisitos: ver requirements.txt (pywin32 + playwright, Windows + Microsoft
Excel instalado). Configuração: ver config.json (criado automaticamente com
valores padrão na primeira execução, a partir de config.example.json).
==============================================================================
"""
from amigo.menu import main

if __name__ == "__main__":
    main()
