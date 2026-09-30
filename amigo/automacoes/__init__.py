# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
Pacote das automações SIGEF. Cada módulo aqui dentro implementa uma etapa
do fluxo (CE -> NL -> PP -> Raspar contas -> OB); este `__init__.py` só
reúne todas elas no dicionário `PROGRAMAS`, usado pelo submenu "Executar
automação" (`amigo.menu.menu_programas`) para listar e disparar a opção
escolhida pelo usuário.

Para adicionar uma nova automação no futuro, basta:
    1. Criar "def minha_automacao(dados, config=None, worksheet=None)"
       num módulo deste pacote (novo ou existente).
    2. Importá-la aqui embaixo.
    3. Adicionar uma entrada nova em PROGRAMAS.
A opção já aparece automaticamente no menu.

OBS: a numeração é contígua de 1 a 10 - a "Raspar contas" entrou como 7 e
empurrou "Gerar OB" e "Raspar OB" para 8 e 9, ocupando o slot que estava
reservado; "Confirmar OB" seguiu em 10 (número 9 originalmente reservado).
"""
from .ce import ce, raspar_ce
from .nl import nl, raspar_nl
from .pp import pp, raspar_pp
from .raspar_contas import raspar_conta
from .ob import gerar_ob, raspar_ob, confirmar_ob

PROGRAMAS = {
    1: ("CE", ce),
    2: ("Raspar CE", raspar_ce),
    3: ("NL", nl),
    4: ("Raspar NL", raspar_nl),
    5: ("PP", pp),
    6: ("Raspar PP", raspar_pp),
    7: ("Raspar contas", raspar_conta),
    8: ("Gerar OB", gerar_ob),
    9: ("Raspar OB", raspar_ob),
    10: ("Confirmar OB", confirmar_ob),
}

__all__ = [
    "PROGRAMAS",
    "ce", "raspar_ce", "nl", "raspar_nl", "pp", "raspar_pp",
    "raspar_conta", "gerar_ob", "raspar_ob", "confirmar_ob",
]
