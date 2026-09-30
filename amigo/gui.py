# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
INTERFACE GRÁFICA (Tkinter)
=============================
Aplicativo com janela, para quem prefere não usar o terminal. Reaproveita
TODA a lógica já existente (amigo.automacoes, amigo.excel, amigo.colunas,
amigo.config) sem alterar nenhuma delas - a GUI só chama essas mesmas
funções a partir de botões e formulários, em vez de input() no terminal.

Tema visual: preto + azul escuro, com fontes maiores que o padrão do
Windows e instruções em português simples em cada aba - pensado para
quem não está acostumado a usar computador (ex: pessoas idosas).

Ponto de entrada: `main_gui.py`, na raiz do projeto, chama `iniciar()`.
"""
import ctypes
import os
import queue
import sys
import threading
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from . import __version__
from . import execucao
from .automacoes import PROGRAMAS
from .colunas import (
    CAMPOS_COLUNAS_AVANCADAS,
    CAMPOS_COLUNAS_BASICAS,
    CAMPOS_COLUNAS_ENCADEADAS,
    COLUNAS_PADRAO,
    NOMES_CURTOS_COLUNAS,
    _detectar_colisoes,
    aplicar_colunas,
)
from .config import carregar_configuracoes
from .constantes import ano_do_exercicio
from .config_io import salvar_configuracoes
from .estatisticas import (
    CATEGORIAS_ESTATISTICAS,
    calcular_estatisticas,
    obter_tempos_manuais,
    salvar_tempos_manuais,
)
from .excel import conectar_planilha, ler_dados
from .lancamento_manual import (
    ETAPAS,
    NOMES_ETAPAS,
    executar_cadeia,
    montar_item,
    padronizar_campo,
    CHAVE_CONFIG_PARAMETROS_MANUAL,
    aviso_mes_referencia,
    config_para_lancamento_manual,
    montar_parametros_manuais,
    padronizar_parametro_manual,
    parametros_manuais_salvos,
    texto_resultados,
)
from .log import adicionar_ouvinte, log_aviso, log_erro, log_info, log_sucesso, remover_ouvinte
from .observacoes import (
    CHAVES_CONFIG_OBSERVACAO,
    LIMITE_AVISO_CARACTERES,
    NOMES_OBSERVACAO,
    TEXTO_PADRAO_OBSERVACAO,
    TIPOS_OBSERVACAO,
    aplicar_marcacoes,
)
from .relatorio import (
    CATEGORIAS_PRINCIPAIS,
    formatar_duracao,
    meses_disponiveis,
    registrar_execucao,
    resumo_mes,
)
from .setup_dependencias import instalar_dependencias, verificar_pendencias

# Agrupamento dos botões de automação na aba "Automações" - mesma ordem e
# nomes de amigo.automacoes.PROGRAMAS, só organizados por etapa da cadeia
# CE -> NL -> PP -> Raspar contas -> OB.
GRUPOS_AUTOMACOES = [
    ("CE - Despesa Certificada", [1, 2]),
    ("NL - Nota de Lançamento", [3, 4]),
    ("PP - Preparação de Pagamento", [5, 6]),
    ("Conferir contas (Raspar contas)", [7]),
    ("OB - Ordem Bancária", [8, 9, 10]),
]

# Frase curta embaixo do título de cada etapa (o que ela faz, em
# português simples) - na mesma ordem de GRUPOS_AUTOMACOES.
DESCRICOES_GRUPOS_AUTOMACOES = {
    "CE - Despesa Certificada": "Lança a despesa de cada pessoa no SIGEF e anota o número da CE na planilha.",
    "NL - Nota de Lançamento": "Liquida a despesa (usa a CE já anotada) e anota o número da NL.",
    "PP - Preparação de Pagamento": "Prepara o pagamento (usa a CE e a NL) conferindo banco, agência e conta.",
    "Conferir contas (Raspar contas)": "Confere se o banco, a agência e a conta do SIGEF são os mesmos da planilha.",
    "OB - Ordem Bancária": "Gera a Ordem Bancária, juntando até 30 pessoas em cada OB.",
}

# Texto do botão (com ícone) e a explicação que aparece ao passar o mouse,
# por número de automação (mesmos números de amigo.automacoes.PROGRAMAS -
# o NOME da automação, usado no Relatório, continua o mesmo).
BOTOES_AUTOMACOES = {
    1: ("▶  Fazer CE", "Lança uma CE nova no SIGEF para cada linha da planilha e anota o número dela."),
    2: ("🔎  Buscar CE já feita", "Use quando a CE JÁ foi lançada antes: o programa só procura o número dela no SIGEF e anota na planilha. Não lança nada novo."),
    3: ("▶  Fazer NL", "Liquida cada despesa usando a CE anotada na planilha, e anota o número da NL."),
    4: ("🔎  Buscar NL já feita", "Use quando a NL JÁ foi gerada: só procura o número dela no SIGEF e anota na planilha."),
    5: ("▶  Fazer PP", "Prepara o pagamento usando a CE e a NL da planilha. Se o SIGEF recusar, o motivo é anotado no lugar do número."),
    6: ("🔎  Buscar PP já feita", "Use quando a PP JÁ foi gerada: só procura o número dela e confere CE, NL e valor."),
    7: ("🔎  Conferir contas", "Confere banco, agência e conta gravados no SIGEF e escreve 'Igual' ou 'Diferente' na planilha."),
    8: ("▶  Gerar OB", "Gera as Ordens Bancárias em lotes de até 30 pessoas e anota o número da OB de cada uma."),
    9: ("⏳  Buscar OB (em breve)", "Ainda não está pronta - clicar não faz nada no SIGEF."),
    10: ("⏳  Confirmar OB (em breve)", "Ainda não está pronta - clicar não faz nada no SIGEF."),
}

# Campos da aba "Lançamento Manual": (campo, título na tela, exemplo que
# aparece ao lado enquanto o campo está vazio). A ordem é a ordem na tela.
CAMPOS_TELA_MANUAL = [
    ("cpf", "👤  CPF de quem vai receber", "Ex: 529.982.247-25 - pode digitar só os números"),
    ("ne", "📄  Nota de Empenho (NE)", "Ex: 2026NE001234 - ou só o número: 1234"),
    ("valor", "💰  Valor a pagar", "Ex: 1400 = R$ 1.400,00  •  1400,50 = R$ 1.400,50"),
    ("banco", "🏦  Banco", "Ex: 1 = Banco do Brasil - só números"),
    ("agencia", "🏢  Agência", "Ex: 1178-9 - com ou sem o traço"),
    ("conta", "💳  Conta", "Ex: 74.508-1 - com ou sem ponto e traço"),
]
TITULOS_CAMPOS_MANUAL = {campo: titulo for campo, titulo, _dica in CAMPOS_TELA_MANUAL}

# Parâmetros próprios da aba "Lançamento Manual" - os mesmos 3 da aba
# "Parâmetros" (mesmos exemplos), mas guardados separados.
CAMPOS_PARAMETROS_MANUAL = [
    ("data", "📅  Data", "Ex: 30062026 ou 30/06/2026"),
    ("processo", "📁  Processo", "Processo do Amigo Voluntário. Ex: 0029.037004/2026-72"),
    ("mes_referencia", "🗓  Mês referência", "Ex: 06/2026 - sempre o mês anterior"),
]
TITULOS_PARAMETROS_MANUAL = {chave: titulo for chave, titulo, _dica in CAMPOS_PARAMETROS_MANUAL}

# ----------------------------------------------------------------------
# TEMA VISUAL (interface nova): escuro, calmo e com bastante respiro.
#
# Camadas de cor, da mais funda para a mais clara:
#   menu lateral (COR_LATERAL) -> fundo das telas (COR_FUNDO) -> cartões
#   (COR_CARTAO) -> caixas de digitar (COR_CAMPO, mais escuras que o
#   cartão, para "afundar" e mostrar onde se escreve).
# Um único azul forte (COR_AZUL_CLARO) marca a ação principal de cada tela;
# o resto dos botões é discreto (COR_BOTAO) para não disputar atenção.
#
# "clam" é o único tema ttk que aceita customização completa de cores no
# Windows - os temas nativos ("vista"/"xpnative") ignoram a maior parte
# das cores configuradas abaixo.
# ----------------------------------------------------------------------
# Escala da tela (1.0 = 100% no Windows; 1.25 = 125%; 1.5 = 150%...).
# Calculada em `_aplicar_tema_escuro`, quando a janela já existe. Tudo que é
# medido em PIXELS (margens, largura do menu, altura das linhas das
# tabelas) passa por `px()`, para crescer junto com as letras - senão, com
# o Windows em 125%/150%, as letras cresceriam e os espaços não, e a tela
# ficaria apertada.
ESCALA = 1.0


def px(valor):
    """Converte pixels "de 100%" para a escala real da tela. Aceita um
    número ou uma tupla de números (ex: padding)."""
    if isinstance(valor, (tuple, list)):
        return tuple(px(v) for v in valor)
    return int(round(valor * ESCALA))


COR_LATERAL = "#0a0e13"            # menu lateral (o mais escuro)
COR_FUNDO = "#0f141a"              # fundo das telas
COR_CARTAO = "#161d26"             # cartões (blocos de conteúdo)
COR_CAMPO = "#0c1117"              # caixas de digitar e tabelas
COR_BORDA = "#27313d"              # bordas sutis
COR_PAINEL_CLARO = "#1d2631"       # destaque suave (fundo da trilha da barra, botão desligado)
COR_BOTAO = "#233041"              # botão comum
COR_BOTAO_HOVER = "#2e3e53"        # botão comum com o mouse em cima
COR_MENU_ATIVO = "#1a2940"         # item do menu lateral selecionado
COR_AZUL = "#1f4e8c"               # azul escuro (seleção em tabelas/listas)
COR_AZUL_CLARO = "#4c9aff"         # azul vivo: ação principal, títulos de destaque, foco
COR_AZUL_HOVER = "#6aabff"         # azul principal com o mouse em cima
COR_TEXTO = "#eef2f7"              # texto principal
COR_TEXTO_SECUNDARIO = "#9aa9ba"   # texto de apoio (explicações, dicas)
COR_TEXTO_APAGADO = "#66768a"      # títulos de grupo do menu, textos bem secundários
COR_VERDE = "#3ddc84"              # deu certo
COR_VERMELHO = "#ff6b6b"           # erro / corrigir
COR_AMARELO = "#ffcc66"            # atenção
COR_SIMULACAO = "#3a2e05"          # fundo da faixa "modo simulação ligado"

FAMILIA_FONTE = "Segoe UI"
FONTE_TITULO = (FAMILIA_FONTE, 22, "bold")
FONTE_SECAO = (FAMILIA_FONTE, 15, "bold")
FONTE_SUBTITULO = (FAMILIA_FONTE, 13, "bold")
FONTE_TEXTO = (FAMILIA_FONTE, 12)
FONTE_TEXTO_NEGRITO = (FAMILIA_FONTE, 12, "bold")
FONTE_BOTAO = (FAMILIA_FONTE, 12, "bold")
FONTE_BOTAO_PRINCIPAL = (FAMILIA_FONTE, 13, "bold")
FONTE_MENU = (FAMILIA_FONTE, 13)
FONTE_MENU_ATIVO = (FAMILIA_FONTE, 13, "bold")
FONTE_GRUPO_MENU = (FAMILIA_FONTE, 10, "bold")
FONTE_NUMERO_GRANDE = (FAMILIA_FONTE, 24, "bold")
FONTE_LOG = ("Consolas", 11)
FONTE_DICA = (FAMILIA_FONTE, 11)
FONTE_DICA_NEGRITO = (FAMILIA_FONTE, 11, "bold")

ESTILO_ENTRADA_NORMAL = "TEntry"
ESTILO_ENTRADA_CONFLITO = "Conflito.TEntry"

# Cor/texto usados para colorir cada linha do painel de mensagens (log) de
# acordo com o nível (o mesmo "nivel" que amigo/log.py passa para os
# ouvintes).
PREFIXOS_NIVEL_LOG = {
    "sucesso": "✅",
    "erro": "❗",
    "aviso": "⚠",
    "info": "•",
}
CORES_NIVEL_LOG = {
    "sucesso": COR_VERDE,
    "erro": COR_VERMELHO,
    "aviso": COR_AMARELO,
    "info": COR_TEXTO,
}

# Estilos de rótulo/quadro que ganham uma "cópia" com fundo de cartão
# ("Cartao.<estilo>") - usada automaticamente para tudo que estiver dentro
# de um `Cartao` (ver `_ajustar_cartoes` e `_estilo`).
ESTILOS_COM_VERSAO_CARTAO = [
    "TLabel", "TFrame", "Titulo.TLabel", "Subtitulo.TLabel", "Instrucao.TLabel",
    "Sucesso.TLabel", "Erro.TLabel", "Aviso.TLabel", "Dica.TLabel", "DicaOk.TLabel",
    "DicaAviso.TLabel", "DicaErro.TLabel", "Campo.TLabel", "Secao.TLabel",
    "NumeroGrande.TLabel", "NumeroGrandeOk.TLabel", "Link.TButton",
    "TRadiobutton", "TCheckbutton",
]


def _aplicar_tema_escuro(root: tk.Tk) -> None:
    """
    Configura o tema (cores, fontes e espaçamentos) usado em todo o
    aplicativo. Chamada uma única vez, em `iniciar()`, antes de montar
    qualquer tela.
    """
    global ESCALA
    try:
        ESCALA = max(1.0, min(3.0, root.winfo_fpixels("1i") / 96.0))
    except tk.TclError:
        ESCALA = 1.0

    root.configure(bg=COR_FUNDO)
    root.option_add("*Font", FONTE_TEXTO)
    # O "*Font" acima vale para caixas de texto/listas, mas também entrava
    # no rótulo (ttk.Label) e ESCONDIA a fonte dos estilos: títulos e
    # negritos saíam todos iguais. Fonte vazia = usa a fonte do estilo.
    root.option_add("*TLabel.Font", "")

    estilo = ttk.Style(root)
    estilo.theme_use("clam")

    estilo.configure(".", background=COR_FUNDO, foreground=COR_TEXTO, font=FONTE_TEXTO)
    estilo.configure("TFrame", background=COR_FUNDO)
    estilo.configure("TLabel", background=COR_FUNDO, foreground=COR_TEXTO, font=FONTE_TEXTO)
    estilo.configure("Titulo.TLabel", foreground=COR_TEXTO, font=FONTE_TITULO)
    estilo.configure("Secao.TLabel", foreground=COR_TEXTO, font=FONTE_SECAO)
    estilo.configure("Subtitulo.TLabel", foreground=COR_TEXTO, font=FONTE_SUBTITULO)
    estilo.configure("Instrucao.TLabel", foreground=COR_TEXTO_SECUNDARIO, font=FONTE_TEXTO)
    estilo.configure("Sucesso.TLabel", foreground=COR_VERDE, font=FONTE_TEXTO_NEGRITO)
    estilo.configure("Erro.TLabel", foreground=COR_VERMELHO, font=FONTE_TEXTO_NEGRITO)
    estilo.configure("Aviso.TLabel", foreground=COR_AMARELO, font=FONTE_TEXTO_NEGRITO)
    estilo.configure("Dica.TLabel", foreground=COR_TEXTO_SECUNDARIO, font=FONTE_DICA)
    estilo.configure("DicaOk.TLabel", foreground=COR_VERDE, font=FONTE_DICA_NEGRITO)
    estilo.configure("DicaAviso.TLabel", foreground=COR_AMARELO, font=FONTE_DICA_NEGRITO)
    estilo.configure("DicaErro.TLabel", foreground=COR_VERMELHO, font=FONTE_DICA_NEGRITO)
    estilo.configure("Campo.TLabel", foreground=COR_TEXTO, font=FONTE_TEXTO_NEGRITO)
    estilo.configure("NumeroGrande.TLabel", foreground=COR_TEXTO, font=FONTE_NUMERO_GRANDE)
    estilo.configure("NumeroGrandeOk.TLabel", foreground=COR_VERDE, font=FONTE_NUMERO_GRANDE)

    # ---- Opções (bolinha / caixinha de marcar) ----
    for nome_opcao in ("TRadiobutton", "TCheckbutton"):
        estilo.configure(
            nome_opcao, background=COR_FUNDO, foreground=COR_TEXTO, font=FONTE_TEXTO,
            indicatorbackground=COR_CAMPO, indicatorforeground=COR_AZUL_CLARO,
            upperbordercolor=COR_BORDA, lowerbordercolor=COR_BORDA, focuscolor=COR_FUNDO,
            padding=px((2, 6)), indicatormargin=px((0, 0, 10, 0)),
        )
        estilo.map(
            nome_opcao,
            background=[("active", COR_FUNDO)],
            indicatorbackground=[("selected", COR_AZUL_CLARO), ("pressed", COR_BOTAO)],
            foreground=[("disabled", COR_TEXTO_APAGADO)],
        )
    estilo.configure(
        "Lateral.TCheckbutton", background=COR_LATERAL, foreground=COR_TEXTO, font=FONTE_TEXTO_NEGRITO,
        indicatorbackground=COR_CAMPO, indicatorforeground="#051226", focuscolor=COR_LATERAL,
        upperbordercolor=COR_BORDA, lowerbordercolor=COR_BORDA, padding=px((2, 6)),
        indicatormargin=px((0, 0, 10, 0)),
    )
    estilo.map(
        "Lateral.TCheckbutton", background=[("active", COR_LATERAL)],
        indicatorbackground=[("selected", COR_AMARELO)],
        foreground=[("disabled", COR_TEXTO_APAGADO)],
    )

    # ---- Faixa do modo simulação (topo da área principal) ----
    estilo.configure("Simulacao.TFrame", background=COR_SIMULACAO)
    estilo.configure("Simulacao.TLabel", background=COR_SIMULACAO, foreground=COR_AMARELO, font=FONTE_TEXTO_NEGRITO)
    estilo.configure(
        "Simulacao.TButton", background="#6b5200", foreground=COR_TEXTO, font=FONTE_BOTAO,
        padding=px((14, 6)), borderwidth=0,
    )
    estilo.map("Simulacao.TButton", background=[("active", "#826400")])

    # ---- Cartões ----
    estilo.configure(
        "Cartao.TFrame", background=COR_CARTAO, bordercolor=COR_BORDA,
        lightcolor=COR_CARTAO, darkcolor=COR_CARTAO, borderwidth=1, relief="solid",
    )
    estilo.configure("CartaoTitulo.TLabel", background=COR_CARTAO, foreground=COR_TEXTO, font=FONTE_SECAO)
    for nome in ESTILOS_COM_VERSAO_CARTAO:
        if nome == "TFrame":
            estilo.configure("Cartao.Interno.TFrame", background=COR_CARTAO)
        else:
            estilo.configure(f"Cartao.{nome}", background=COR_CARTAO)

    # ---- Painel "Como usar" ----
    estilo.configure("Ajuda.TFrame", background=COR_CARTAO)
    estilo.configure("Ajuda.TLabel", background=COR_CARTAO, foreground=COR_TEXTO, font=FONTE_TEXTO)
    estilo.configure("AjudaNumero.TLabel", background=COR_CARTAO, foreground=COR_AZUL_CLARO, font=FONTE_SECAO)

    # ---- Menu lateral ----
    estilo.configure("Lateral.TFrame", background=COR_LATERAL)
    estilo.configure("LateralMarca.TLabel", background=COR_LATERAL, foreground=COR_TEXTO, font=(FAMILIA_FONTE, 20, "bold"))
    estilo.configure("LateralSub.TLabel", background=COR_LATERAL, foreground=COR_TEXTO_SECUNDARIO, font=FONTE_DICA)
    estilo.configure("LateralGrupo.TLabel", background=COR_LATERAL, foreground=COR_TEXTO_APAGADO, font=FONTE_GRUPO_MENU)
    estilo.configure("LateralDireitos.TLabel", background=COR_LATERAL, foreground=COR_TEXTO_APAGADO, font=(FAMILIA_FONTE, 9))
    estilo.configure("LateralStatus.TLabel", background=COR_LATERAL, foreground=COR_TEXTO_SECUNDARIO, font=FONTE_DICA)
    estilo.configure("LateralStatusOk.TLabel", background=COR_LATERAL, foreground=COR_VERDE, font=FONTE_DICA_NEGRITO)
    estilo.configure("LateralStatusAviso.TLabel", background=COR_LATERAL, foreground=COR_AMARELO, font=FONTE_DICA_NEGRITO)
    estilo.configure(
        "Menu.TButton", background=COR_LATERAL, foreground=COR_TEXTO_SECUNDARIO, font=FONTE_MENU,
        anchor="w", padding=px((18, 12)), borderwidth=0, focuscolor=COR_LATERAL,
    )
    estilo.map("Menu.TButton", background=[("active", COR_BOTAO)], foreground=[("active", COR_TEXTO)])
    estilo.configure(
        "MenuAtivo.TButton", background=COR_MENU_ATIVO, foreground=COR_TEXTO, font=FONTE_MENU_ATIVO,
        anchor="w", padding=px((18, 12)), borderwidth=0, focuscolor=COR_MENU_ATIVO,
    )
    estilo.map("MenuAtivo.TButton", background=[("active", COR_MENU_ATIVO)])
    estilo.configure("MenuMarca.TFrame", background=COR_LATERAL)
    estilo.configure("MenuMarcaAtiva.TFrame", background=COR_AZUL_CLARO)

    # ---- Gaveta de mensagens ----
    estilo.configure("Gaveta.TFrame", background=COR_LATERAL)
    estilo.configure("Gaveta.TLabel", background=COR_LATERAL, foreground=COR_TEXTO_SECUNDARIO, font=FONTE_DICA)
    for nivel, cor in CORES_NIVEL_LOG.items():
        estilo.configure(f"Gaveta{nivel.capitalize()}.TLabel", background=COR_LATERAL, foreground=cor, font=FONTE_DICA)
    estilo.configure(
        "Gaveta.TButton", background=COR_LATERAL, foreground=COR_AZUL_CLARO, font=FONTE_DICA_NEGRITO,
        padding=px((10, 6)), borderwidth=0, focuscolor=COR_LATERAL,
    )
    estilo.map("Gaveta.TButton", background=[("active", COR_BOTAO)])

    # ---- Botões ----
    estilo.configure(
        "TButton", background=COR_BOTAO, foreground=COR_TEXTO, font=FONTE_BOTAO,
        padding=px((16, 10)), borderwidth=0, focuscolor=COR_AZUL_CLARO,
    )
    estilo.map(
        "TButton",
        background=[("disabled", COR_PAINEL_CLARO), ("active", COR_BOTAO_HOVER)],
        foreground=[("disabled", COR_TEXTO_APAGADO)],
    )
    # Ação principal de cada tela - a única em azul forte.
    estilo.configure(
        "Principal.TButton", background=COR_AZUL_CLARO, foreground="#051226",
        font=FONTE_BOTAO_PRINCIPAL, padding=px((22, 12)), borderwidth=0,
    )
    estilo.map(
        "Principal.TButton",
        background=[("disabled", COR_PAINEL_CLARO), ("active", COR_AZUL_HOVER)],
        foreground=[("disabled", COR_TEXTO_APAGADO)],
    )
    # Botão com cara de link (abre/fecha o "Como usar").
    estilo.configure(
        "Link.TButton", background=COR_FUNDO, foreground=COR_AZUL_CLARO, font=FONTE_TEXTO_NEGRITO,
        padding=px((10, 6)), borderwidth=0, focuscolor=COR_FUNDO,
    )
    estilo.map("Link.TButton", background=[("active", COR_BOTAO)], foreground=[("active", COR_TEXTO)])

    # ---- Caixas de digitar ----
    estilo.configure(
        "TEntry", fieldbackground=COR_CAMPO, foreground=COR_TEXTO,
        insertcolor=COR_TEXTO, bordercolor=COR_BORDA, lightcolor=COR_BORDA, darkcolor=COR_BORDA,
        borderwidth=1, padding=px(10), font=FONTE_TEXTO,
    )
    estilo.map(
        "TEntry",
        fieldbackground=[("disabled", COR_PAINEL_CLARO)],
        bordercolor=[("focus", COR_AZUL_CLARO)],
        lightcolor=[("focus", COR_AZUL_CLARO)],
    )
    # Coluna em conflito (duas linhas usando a mesma letra).
    estilo.configure(ESTILO_ENTRADA_CONFLITO, fieldbackground="#4a1f24", foreground="#ffd9d9")

    # Combobox (seletor de mês) - a lista suspensa é uma Listbox "crua" do
    # Tk, sem tema próprio, por isso as cores dela vão via `option_add`.
    estilo.configure(
        "TCombobox", fieldbackground=COR_CAMPO, background=COR_BOTAO,
        foreground=COR_TEXTO, arrowcolor=COR_TEXTO, bordercolor=COR_BORDA,
        lightcolor=COR_BORDA, darkcolor=COR_BORDA, padding=px(8), font=FONTE_TEXTO,
    )
    estilo.map(
        "TCombobox",
        fieldbackground=[("readonly", COR_CAMPO), ("disabled", COR_PAINEL_CLARO)],
        foreground=[("disabled", COR_TEXTO_SECUNDARIO)],
    )
    root.option_add("*TCombobox*Listbox.background", COR_CAMPO)
    root.option_add("*TCombobox*Listbox.foreground", COR_TEXTO)
    root.option_add("*TCombobox*Listbox.selectBackground", COR_AZUL)
    root.option_add("*TCombobox*Listbox.selectForeground", COR_TEXTO)
    root.option_add("*TCombobox*Listbox.font", FONTE_TEXTO)

    for nome_estilo_barra in ("TProgressbar", "Horizontal.TProgressbar"):
        estilo.configure(
            nome_estilo_barra, background=COR_AZUL_CLARO, troughcolor=COR_PAINEL_CLARO,
            bordercolor=COR_PAINEL_CLARO, lightcolor=COR_AZUL_CLARO, darkcolor=COR_AZUL_CLARO,
            thickness=px(10),
        )

    # Tabelas (Treeview) - linhas altas, sem bordas pesadas.
    estilo.configure(
        "Treeview", background=COR_CAMPO, fieldbackground=COR_CAMPO, foreground=COR_TEXTO,
        font=FONTE_TEXTO, rowheight=px(36), bordercolor=COR_BORDA, lightcolor=COR_BORDA,
        darkcolor=COR_BORDA, borderwidth=1,
    )
    estilo.map("Treeview", background=[("selected", COR_AZUL)], foreground=[("selected", COR_TEXTO)])
    estilo.configure(
        "Treeview.Heading", background=COR_CARTAO, foreground=COR_TEXTO_SECUNDARIO,
        font=FONTE_DICA_NEGRITO, relief="flat", padding=px((8, 8)), borderwidth=0,
    )
    estilo.map("Treeview.Heading", background=[("active", COR_BOTAO)])

    for nome_estilo_scroll in ("TScrollbar", "Vertical.TScrollbar"):
        estilo.configure(
            nome_estilo_scroll, background=COR_BOTAO, troughcolor=COR_FUNDO,
            bordercolor=COR_FUNDO, arrowcolor=COR_TEXTO_SECUNDARIO, lightcolor=COR_BOTAO,
            darkcolor=COR_BOTAO, gripcount=0,
        )
        estilo.map(nome_estilo_scroll, background=[("active", COR_BOTAO_HOVER)])


def _estilizar_caixa_texto(caixa: tk.Text) -> None:
    """Aplica o tema a um `tk.Text` (widget "cru" do Tk, que não segue o
    `ttk.Style` automaticamente - precisa ser configurado à mão)."""
    caixa.configure(
        background=COR_CAMPO, foreground=COR_TEXTO, insertbackground=COR_TEXTO,
        selectbackground=COR_AZUL, selectforeground=COR_TEXTO,
        font=FONTE_LOG, relief="flat", padx=px(14), pady=px(10),
        highlightthickness=0, borderwidth=0,
    )


def _criar_area_rolavel(aba_pai, fundo: str = None, margens=(40, 28, 40, 32), estilo_conteudo: str = "TFrame"):
    """
    Cria, dentro de `aba_pai`, uma área com rolagem vertical e devolve o
    quadro interno onde o conteúdo deve ser montado - o MESMO esquema em
    todas as telas (e no menu lateral), para nada ficar cortado numa janela
    menor.

    - A roda do mouse funciona em QUALQUER ponto da área (em cima de
      cartões, textos, botões...), não só no fundo: quem trata a roda é um
      único "ouvinte" do programa (`AplicativoAmigo._rolar_com_mouse`), que
      descobre qual área está embaixo do mouse. Antes a roda só era ligada
      quando o mouse estava sobre o fundo vazio - e, como os cartões cobrem
      quase tudo, a rolagem parecia não funcionar.
    - A barra de rolagem só aparece quando o conteúdo não cabe na janela.
    """
    fundo = fundo or COR_FUNDO
    canvas = tk.Canvas(aba_pai, highlightthickness=0, background=fundo, borderwidth=0,
                       yscrollincrement=px(60))
    canvas._area_rolavel = True
    rolagem = ttk.Scrollbar(aba_pai, orient="vertical", command=canvas.yview)
    conteudo = ttk.Frame(canvas, padding=px(margens), style=estilo_conteudo)
    janela = canvas.create_window((0, 0), window=conteudo, anchor="nw")
    canvas.configure(yscrollcommand=rolagem.set)
    canvas.pack(side="left", fill="both", expand=True)

    def ajustar(_evento=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        cabe = conteudo.winfo_reqheight() <= canvas.winfo_height()
        if cabe:
            rolagem.pack_forget()
            canvas.yview_moveto(0)
        elif not rolagem.winfo_ismapped():
            rolagem.pack(side="right", fill="y", before=canvas)

    def ao_redimensionar(evento):
        canvas.itemconfigure(janela, width=evento.width)
        ajustar()

    conteudo.bind("<Configure>", ajustar, add="+")
    canvas.bind("<Configure>", ao_redimensionar, add="+")
    return conteudo


# ----------------------------------------------------------------------
# COMPONENTES DE INTERFACE REAPROVEITADOS EM TODAS AS TELAS
#
# Ideia geral (UI/UX para quem não tem prática com computador):
#   - cada tela mostra SEMPRE só o essencial: título, 1 frase do que ela
#     faz e os cartões com o conteúdo;
#   - o passo a passo completo fica no botão "❓ Como usar", que abre e
#     fecha;
#   - explicações longas aparecem ao parar o mouse em cima de um botão
#     (`Dica`);
#   - cada bloco de conteúdo é um `Cartao`, com bastante espaço interno.
# ----------------------------------------------------------------------
NUMEROS_CIRCULADOS = "①②③④⑤⑥⑦⑧⑨⑩"


class Cartao(ttk.Frame):
    """
    Bloco de conteúdo com fundo próprio, borda sutil e título dentro -
    substitui as antigas caixas com borda (LabelFrame), que deixavam a tela
    com cara de "apertada". Aceita os mesmos `text=` e `padding=` que o
    LabelFrame aceitava, e os filhos podem usar pack OU grid normalmente:
    o título é posicionado com `place`, que não disputa espaço com eles.
    """

    MARGEM = 24
    ALTURA_TITULO = 44

    def __init__(self, master, text: str = "", padding=None, **kwargs):
        margem = px(self.MARGEM)
        topo = margem + (px(self.ALTURA_TITULO) if text else 0)
        super().__init__(
            master, style="Cartao.TFrame",
            padding=(margem, topo, margem, margem), **kwargs,
        )
        self.rotulo_titulo = None
        if text:
            self.rotulo_titulo = ttk.Label(self, text=text, style="CartaoTitulo.TLabel")
            self.rotulo_titulo.place(x=margem, y=margem - px(4), bordermode="ignore")


def _dentro_de_cartao(widget) -> bool:
    atual = widget
    while atual is not None:
        if isinstance(atual, Cartao):
            return True
        atual = getattr(atual, "master", None)
    return False


def _configurar(widget, style=None, **opcoes) -> None:
    """`widget.config(**opcoes)` + troca de estilo pelo `_estilo` (que
    respeita o fundo do cartão)."""
    if opcoes:
        widget.config(**opcoes)
    if style:
        _estilo(widget, style)


def _estilo(widget, estilo: str) -> None:
    """
    Troca o estilo de um rótulo levando em conta ONDE ele está: dentro de
    um cartão usa a versão "Cartao.<estilo>" (mesma cor de letra, fundo do
    cartão). Use sempre isto em vez de `widget.config(style=...)` para
    rótulos que mudam de cor (ok/erro/aviso) depois de criados.
    """
    if _dentro_de_cartao(widget) and not estilo.startswith("Cartao"):
        estilo = "Cartao.Interno.TFrame" if estilo == "TFrame" else f"Cartao.{estilo}"
    widget.configure(style=estilo)


def _ajustar_cartoes(raiz) -> None:
    """
    Percorre a tela a partir de `raiz` e deixa tudo que está dentro de um
    `Cartao` com o fundo do cartão (rótulos, quadros e botões-link), além
    de ligar a "quebra de linha automática" nos textos longos - eles
    passam a quebrar na largura real da janela, em vez de num número fixo
    de pixels (que cortava o texto em janelas menores). Chamada depois de
    montar as telas e depois de redesenhar partes dinâmicas.
    """
    pilha = [raiz]
    while pilha:
        widget = pilha.pop()
        pilha.extend(widget.winfo_children())
        classe = widget.winfo_class()
        if classe in ("TLabel", "TFrame", "TButton", "TRadiobutton", "TCheckbutton") and not isinstance(widget, Cartao):
            estilo_atual = str(widget.cget("style")) or classe
            if _dentro_de_cartao(widget) and not estilo_atual.startswith("Cartao"):
                if classe == "TFrame" and estilo_atual == "TFrame":
                    widget.configure(style="Cartao.Interno.TFrame")
                elif estilo_atual in ESTILOS_COM_VERSAO_CARTAO:
                    widget.configure(style=f"Cartao.{estilo_atual}")
        if classe == "TLabel" and not getattr(widget, "_quebra_automatica", False):
            try:
                largura = int(float(str(widget.cget("wraplength")) or 0))
            except ValueError:
                largura = 0
            estilo_rotulo = str(widget.cget("style"))
            if (largura == 0 and len(str(widget.cget("text"))) > 45
                    and not estilo_rotulo.startswith("Gaveta") and not str(widget.cget("width"))):
                largura = 600  # texto longo sem quebra definida: quebra também
                widget.configure(justify="left")
            if largura > 0:
                # Todo texto com quebra de linha passa a quebrar na largura
                # REAL do espaço onde está (cartão, coluna...), e não num
                # número fixo de pixels - que cortava o texto em janelas
                # menores ou com o Windows em 125%/150%.
                widget._quebra_automatica = True
                widget.configure(wraplength=px(largura))
                widget.master.bind(
                    "<Configure>", lambda _evento, rotulo=widget: _requebrar(rotulo), add="+",
                )


def _requebrar(rotulo) -> None:
    """Ajusta a quebra de linha de `rotulo` ao espaço livre no quadro dele:
    largura do quadro, menos as margens internas e o que estiver ao lado
    (ex: o número ① do passo a passo)."""
    if not rotulo.winfo_exists():
        return
    quadro = rotulo.master
    largura = quadro.winfo_width()
    if largura <= 1:
        return
    try:
        bruto = quadro.cget("padding")
        partes = bruto if isinstance(bruto, (tuple, list)) else str(bruto).split()
        margens = [int(float(str(parte))) for parte in partes]
    except (tk.TclError, ValueError):
        margens = []
    if len(margens) == 1:
        margens = margens * 4
    elif len(margens) == 2:
        margens = margens * 2
    elif len(margens) == 3:
        margens = margens + [margens[1]]
    horizontal = (margens[0] + margens[2]) if len(margens) == 4 else 0
    ao_lado = 0
    for irmao in quadro.pack_slaves():
        if irmao is not rotulo:
            try:
                if irmao.pack_info().get("side") in ("left", "right"):
                    ao_lado += irmao.winfo_reqwidth() + px(16)
            except tk.TclError:
                pass
    nova = max(px(120), largura - horizontal - ao_lado - px(12))
    if abs(nova - int(float(str(rotulo.cget("wraplength")) or 0))) > 2:
        rotulo.configure(wraplength=nova)


class Dica:
    """Balãozinho de explicação que aparece ao deixar o mouse parado por
    um instante em cima de um botão/campo, e some ao tirar o mouse."""

    ATRASO_MS = 450

    def __init__(self, widget, texto: str):
        self.widget = widget
        self.texto = texto
        self._agendado = None
        self._janela = None
        widget.bind("<Enter>", self._agendar, add="+")
        widget.bind("<Leave>", self._esconder, add="+")
        widget.bind("<ButtonPress>", self._esconder, add="+")

    def _agendar(self, _evento=None):
        self._cancelar()
        self._agendado = self.widget.after(self.ATRASO_MS, self._mostrar)

    def _cancelar(self):
        if self._agendado is not None:
            self.widget.after_cancel(self._agendado)
            self._agendado = None

    def _mostrar(self):
        self._agendado = None
        if self._janela is not None or not self.widget.winfo_exists():
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self._janela = tk.Toplevel(self.widget)
        self._janela.wm_overrideredirect(True)
        self._janela.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self._janela, text=f"💡  {self.texto}", justify="left", wraplength=px(400),
            background=COR_BOTAO, foreground=COR_TEXTO, font=FONTE_DICA,
            padx=px(14), pady=px(10), highlightthickness=1, highlightbackground=COR_AZUL_CLARO,
        ).pack()

    def _esconder(self, _evento=None):
        self._cancelar()
        if self._janela is not None:
            self._janela.destroy()
            self._janela = None


def _criar_cabecalho_aba(pai, titulo: str, subtitulo: str, passos=None, ajuda_aberta: bool = False):
    """
    Topo padrão de TODAS as telas: título grande + 1 frase curta do que a
    tela faz e, à direita, o botão "❓ Como usar", que abre/fecha um
    cartão com o passo a passo numerado (①, ②, ...).
    """
    topo = ttk.Frame(pai)
    topo.pack(fill="x", pady=(0, 22))

    linha = ttk.Frame(topo)
    linha.pack(fill="x")
    textos = ttk.Frame(linha)
    textos.pack(side="left", fill="x", expand=True)
    ttk.Label(textos, text=titulo, style="Titulo.TLabel").pack(anchor="w")
    ttk.Label(textos, text=subtitulo, style="Instrucao.TLabel", wraplength=820, justify="left").pack(
        anchor="w", pady=(6, 0)
    )
    if not passos:
        return topo

    painel = Cartao(topo, text="Passo a passo")
    for indice, passo in enumerate(passos):
        item = ttk.Frame(painel)
        item.pack(fill="x", pady=5)
        ttk.Label(item, text=NUMEROS_CIRCULADOS[indice], style="AjudaNumero.TLabel").pack(side="left", anchor="n")
        ttk.Label(item, text=passo, style="Ajuda.TLabel", wraplength=860, justify="left").pack(
            side="left", anchor="w", padx=(14, 0)
        )

    estado = {"aberta": ajuda_aberta}
    botao = ttk.Button(linha, style="Link.TButton")
    botao.pack(side="right", anchor="ne", padx=(16, 0))

    def atualizar():
        if estado["aberta"]:
            botao.config(text="✕  Fechar ajuda")
            painel.pack(fill="x", pady=(18, 0))
        else:
            botao.config(text="❓  Como usar")
            painel.pack_forget()

    def alternar():
        estado["aberta"] = not estado["aberta"]
        atualizar()

    botao.config(command=alternar)
    Dica(botao, "Mostra (ou esconde) o passo a passo desta tela.")
    atualizar()
    return topo


def _texto_ajuda_campo(pai, texto: str) -> ttk.Label:
    """Texto pequeno de ajuda, logo ABAIXO de um campo (exemplo, dica ou
    a conferência do que foi digitado)."""
    return ttk.Label(pai, text=texto, style="Dica.TLabel", wraplength=560, justify="left")


def _mostrar_retorno_campo(rotulo, tipo: str, texto: str) -> None:
    """Pinta o texto de ajuda de um campo conforme a situação:
    'dica' (cinza), 'ok' (verde), 'aviso' (amarelo) ou 'erro' (vermelho) -
    sempre com o mesmo ícone na frente, para ficar fácil de reconhecer."""
    icones = {"dica": "💡", "ok": "✅", "aviso": "⚠", "erro": "❗"}
    estilos = {"dica": "Dica.TLabel", "ok": "DicaOk.TLabel", "aviso": "DicaAviso.TLabel", "erro": "DicaErro.TLabel"}
    rotulo.config(text=f"{icones[tipo]}  {texto}")
    _estilo(rotulo, estilos[tipo])


class AplicativoAmigo:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Amigo {__version__} - Automações SIGEF")
        self.root.geometry("1360x880")
        self.root.minsize(1100, 700)

        self.config = None
        self.worksheet = None
        self.dados = None

        self.fila_log = queue.Queue()
        self.vars_parametros = {}
        self.vars_colunas = {}
        self.entradas_colunas = {}
        self.botoes_automacao = {}
        self.cartoes_relatorio = {}
        self.vars_tempos_manuais = {}
        self.cartoes_estatisticas = {}

        # Aba "Lançamento Manual" (sem planilha).
        self.vars_manual = {}
        self.vars_parametros_manual = {}
        self.entradas_parametros_manual = {}
        self.retornos_parametros_manual = {}
        self.entradas_manual = {}
        self.retornos_manual = {}
        self.itens_manual_pendentes = []    # [(número da pessoa, item)] - ainda não lançados
        self.itens_manual_resultados = {}   # {número da pessoa: item} - já passaram pela cadeia
        self.contador_manual = 0
        self.parar_manual = threading.Event()
        self._manual_rodando = False

        # Só 1 automação por vez mexe no navegador do SIGEF - seja a da
        # planilha (aba Automações), seja a do Lançamento Manual.
        self.automacao_em_andamento = False
        self.var_simulacao = tk.BooleanVar(value=False)   # modo simulação (sempre começa desligado)
        self._nome_automacao_atual = ""

        # Navegação (menu lateral) e gaveta de mensagens.
        self.paginas = {}
        self.itens_menu = {}
        self.pagina_atual = None
        self.gaveta_aberta = False
        self.mensagens_nao_vistas = 0

        self._montar_tela_preparo()
        self.root.after(100, self._drenar_fila_log)

    # ------------------------------------------------------------------
    # TELA DE PREPARO (checagem/instalação de dependências)
    #
    # Roda automaticamente assim que o aplicativo abre, ANTES de mostrar a
    # janela principal - o usuário não precisa clicar em nada nem rodar
    # nenhum comando: se faltar algo (pacotes Python, navegador do
    # Playwright), o próprio programa baixa e instala sozinho aqui.
    # ------------------------------------------------------------------
    def _montar_tela_preparo(self):
        self.frame_preparo = ttk.Frame(self.root, padding=48)
        self.frame_preparo.pack(fill="both", expand=True)

        cartao = Cartao(self.frame_preparo)
        cartao.place(relx=0.5, rely=0.45, anchor="center", relwidth=0.62)
        ttk.Label(cartao, text="🤝  Amigo", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(
            cartao, text="Preparando tudo para você...", style="Secao.TLabel",
        ).pack(anchor="w", pady=(10, 4))
        ttk.Label(
            cartao,
            text="Isso só demora na primeira vez que o programa é aberto neste computador. "
                 "Não precisa fazer nada - é só esperar.",
            style="Instrucao.TLabel", wraplength=640, justify="left",
        ).pack(anchor="w", pady=(0, 20))

        self.barra_preparo = ttk.Progressbar(cartao, mode="indeterminate")
        self.barra_preparo.pack(fill="x")
        self.barra_preparo.start(12)

        self.texto_preparo = tk.Text(cartao, height=10, wrap="word", state="disabled")
        _estilizar_caixa_texto(self.texto_preparo)
        self.texto_preparo.pack(fill="both", expand=True, pady=(20, 0))
        _ajustar_cartoes(cartao)

        threading.Thread(target=self._preparar_em_segundo_plano, daemon=True).start()

    def _log_preparo(self, mensagem: str):
        if not mensagem.strip():
            return
        self.fila_log.put(("preparo", mensagem))

    def _preparar_em_segundo_plano(self):
        try:
            pendencias = verificar_pendencias()
            if not pendencias:
                self._log_preparo("Tudo certo - nada precisa ser baixado desta vez.")
                ok = True
            else:
                self._log_preparo("Primeira vez usando o Amigo neste computador - preparando:")
                for item in pendencias:
                    self._log_preparo(f"  - {item}")
                ok = instalar_dependencias(self._log_preparo)
        except Exception as erro:
            ok = False
            self._log_preparo(f"Erro inesperado ao preparar o sistema: {erro}")
        self.fila_log.put(("preparo_fim", ok))

    def _finalizar_preparo(self, ok: bool):
        self.barra_preparo.stop()
        if not ok:
            messagebox.showwarning(
                "Amigo",
                "Não foi possível preparar automaticamente tudo que o sistema precisa "
                "(veja o log). Você pode continuar, mas alguma automação pode falhar até "
                "isso ser resolvido - reabrir o aplicativo tenta de novo.",
            )
        self.frame_preparo.destroy()
        self._montar_janela_principal()

    # ------------------------------------------------------------------
    # JANELA PRINCIPAL
    # ------------------------------------------------------------------
    # Itens do menu lateral: (grupo, chave da tela, ícone + nome, dica).
    MENU_LATERAL = [
        ("", "inicio", "🏠   Início", "Tela inicial: escolha o que você quer fazer."),
        ("TRABALHAR", "automacoes", "🤖   Com planilha", "Rodar as automações lendo a planilha do Excel."),
        ("TRABALHAR", "manual", "✍   Sem planilha", "Digitar os dados de cada pessoa e lançar CE → NL → PP → OB."),
        ("CONFIGURAR", "parametros", "⚙   Parâmetros", "Data, processo, mês e qual planilha usar."),
        ("CONFIGURAR", "colunas", "🧱   Colunas", "Em qual coluna do Excel está cada informação."),
        ("ACOMPANHAR", "estatisticas", "📊   Estatísticas", "Quanto tempo a automação já economizou."),
        ("ACOMPANHAR", "relatorio", "📅   Relatório", "Quantos documentos foram feitos em cada mês."),
    ]

    def _montar_janela_principal(self):
        self.config = carregar_configuracoes()

        # ---- Menu lateral (esquerda) ----
        barra_lateral = ttk.Frame(self.root, style="Lateral.TFrame", width=px(250))
        barra_lateral.pack(side="left", fill="y")
        barra_lateral.pack_propagate(False)

        # Situação atual, sempre à vista no rodapé do menu (fora da rolagem).
        situacao = ttk.Frame(barra_lateral, style="Lateral.TFrame", padding=px((24, 16, 20, 24)))
        situacao.pack(side="bottom", fill="x")
        # Modo simulação: fica no menu (vale para TODAS as telas) e à vista.
        self.caixa_simulacao = ttk.Checkbutton(
            situacao, text="🧪  Modo simulação", variable=self.var_simulacao,
            style="Lateral.TCheckbutton", command=self._ao_mudar_simulacao,
        )
        self.caixa_simulacao.pack(anchor="w", pady=(0, 14))
        Dica(self.caixa_simulacao,
             "Ligado: as automações fazem tudo no SIGEF de verdade (pesquisam credor, empenho, "
             "conta, conferem valores), mas PARAM ANTES de confirmar - nada é gerado. No lugar do "
             "número aparece \"SIMULADO ✓\" (daria certo) ou \"SIMULADO ✗\" com o motivo. "
             "Ótimo para conferir uma planilha ou treinar alguém.")
        ttk.Label(situacao, text="SITUAÇÃO", style="LateralGrupo.TLabel").pack(anchor="w", pady=(0, 8))
        self.rotulo_lateral_planilha = ttk.Label(situacao, text="", style="LateralStatus.TLabel", wraplength=200, justify="left")
        self.rotulo_lateral_planilha.pack(anchor="w", pady=2)
        self.rotulo_lateral_robo = ttk.Label(situacao, text="", style="LateralStatus.TLabel", wraplength=200, justify="left")
        self.rotulo_lateral_robo.pack(anchor="w", pady=2)
        ttk.Label(
            situacao, text="© 2026 Pedro Henrique Carpina Farias Alves\nTodos os direitos reservados",
            style="LateralDireitos.TLabel", justify="left",
        ).pack(anchor="w", pady=(14, 0))

        # Itens do menu numa área com rolagem própria (em telas baixas ou
        # com o Windows em 150%, nem todos cabiam - e não havia como descer).
        lateral = _criar_area_rolavel(barra_lateral, fundo=COR_LATERAL, margens=(0, 0, 0, 0),
                                      estilo_conteudo="Lateral.TFrame")
        marca = ttk.Frame(lateral, style="Lateral.TFrame", padding=px((24, 28, 24, 22)))
        marca.pack(fill="x")
        ttk.Label(marca, text="🤝  Amigo", style="LateralMarca.TLabel").pack(anchor="w")
        ttk.Label(marca, text=f"Automações do SIGEF\nVersão {__version__}", style="LateralSub.TLabel").pack(anchor="w", pady=(2, 0))

        grupo_atual = None
        for grupo, chave, texto, dica in self.MENU_LATERAL:
            if grupo and grupo != grupo_atual:
                ttk.Label(lateral, text=grupo, style="LateralGrupo.TLabel").pack(anchor="w", padx=26, pady=(20, 6))
                grupo_atual = grupo
            linha = ttk.Frame(lateral, style="Lateral.TFrame")
            linha.pack(fill="x", padx=(0, 12), pady=1)
            marca_ativa = ttk.Frame(linha, style="MenuMarca.TFrame", width=px(4))
            marca_ativa.pack(side="left", fill="y")
            botao = ttk.Button(linha, text=texto, style="Menu.TButton", command=lambda c=chave: self._mostrar_pagina(c))
            botao.pack(side="left", fill="x", expand=True)
            Dica(botao, dica)
            self.itens_menu[chave] = (botao, marca_ativa)

        ttk.Frame(lateral, style="Lateral.TFrame", height=px(16)).pack(fill="x")

        # ---- Área principal (direita): telas + gaveta de mensagens ----
        principal = ttk.Frame(self.root)
        principal.pack(side="left", fill="both", expand=True)

        self._montar_gaveta_mensagens(principal)

        area_paginas = ttk.Frame(principal)
        area_paginas.pack(side="top", fill="both", expand=True)

        # Faixa amarela no topo enquanto o modo simulação estiver ligado.
        self.faixa_simulacao = ttk.Frame(principal, style="Simulacao.TFrame", padding=px((28, 10, 20, 10)))
        self._area_paginas = area_paginas
        ttk.Label(
            self.faixa_simulacao,
            text="🧪  MODO SIMULAÇÃO LIGADO - tudo é conferido no SIGEF, mas as automações param "
                 "antes de confirmar. Nada é gerado.",
            style="Simulacao.TLabel", wraplength=900, justify="left",
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(self.faixa_simulacao, text="Desligar", style="Simulacao.TButton",
                   command=self._desligar_simulacao).pack(side="right", padx=(12, 0))
        for _grupo, chave, _texto, _dica in self.MENU_LATERAL:
            self.paginas[chave] = ttk.Frame(area_paginas)

        # Os nomes antigos ("aba_...") continuam valendo: cada tela é
        # montada exatamente como antes, só que dentro de uma página.
        self.aba_inicio = self.paginas["inicio"]
        self.aba_automacoes = self.paginas["automacoes"]
        self.aba_manual = self.paginas["manual"]
        self.aba_parametros = self.paginas["parametros"]
        self.aba_colunas = self.paginas["colunas"]
        self.aba_estatisticas = self.paginas["estatisticas"]
        self.aba_relatorio = self.paginas["relatorio"]

        self._montar_aba_inicio()
        self._montar_aba_automacoes()
        self._montar_aba_manual()
        self._montar_aba_parametros()
        self._montar_aba_colunas()
        self._montar_aba_estatisticas()
        self._montar_aba_relatorio()
        _ajustar_cartoes(self.root)

        adicionar_ouvinte(self._receber_log)
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar)
        # Roda do mouse: 1 ouvinte para o programa inteiro (ver `_rolar_com_mouse`).
        self.root.bind_all("<MouseWheel>", self._rolar_com_mouse)
        self.root.bind_all("<Button-4>", self._rolar_com_mouse)   # Linux
        self.root.bind_all("<Button-5>", self._rolar_com_mouse)   # Linux
        # Progresso "linha X de Y" das automações (vem da thread da automação).
        execucao.adicionar_ouvinte_progresso(lambda atual, total: self.fila_log.put(("progresso", atual, total)))
        self._atualizar_status_lateral()
        self._mostrar_pagina("inicio")
        try:
            self.root.state("zoomed")  # Windows: abre maximizado (mais espaço, letras não apertadas)
        except tk.TclError:
            pass

        log_info("Interface gráfica pronta.")

    def _rolar_com_mouse(self, evento):
        """
        Rola a área que está EMBAIXO do mouse (tela principal ou menu
        lateral), esteja o mouse sobre o fundo, um cartão, um texto ou um
        botão. Tabelas e a caixa de mensagens rolam o próprio conteúdo
        quando têm o que rolar; quando não têm, a roda rola a tela.
        """
        try:
            widget = self.root.winfo_containing(evento.x_root, evento.y_root)
        except (KeyError, tk.TclError):
            return
        if evento.num == 4:
            passos = -1
        elif evento.num == 5:
            passos = 1
        elif evento.delta:
            passos = -1 if evento.delta > 0 else 1
            passos *= max(1, abs(int(evento.delta)) // 120)
        else:
            return

        while widget is not None:
            classe = widget.winfo_class()
            if classe in ("Treeview", "Text", "Listbox"):
                try:
                    inicio, fim = widget.yview()
                except tk.TclError:
                    inicio, fim = 0.0, 1.0
                if inicio > 0.0 or fim < 1.0:
                    return  # ela mesma rola (comportamento padrão do Tk)
            if getattr(widget, "_area_rolavel", False):
                inicio, fim = widget.yview()
                if inicio <= 0.0 and fim >= 1.0:
                    return  # tudo cabe na tela - nada a rolar
                widget.yview_scroll(passos, "units")
                return
            widget = getattr(widget, "master", None)

    def _mostrar_pagina(self, chave: str):
        """Troca a tela mostrada à direita e destaca o item no menu."""
        if self.pagina_atual == chave:
            return
        if self.pagina_atual is not None:
            self.paginas[self.pagina_atual].pack_forget()
        self.paginas[chave].pack(fill="both", expand=True)
        self.pagina_atual = chave
        for outra, (botao, marca_ativa) in self.itens_menu.items():
            ativa = outra == chave
            botao.configure(style="MenuAtivo.TButton" if ativa else "Menu.TButton")
            marca_ativa.configure(style="MenuMarcaAtiva.TFrame" if ativa else "MenuMarca.TFrame")
        # Telas que mostram números sempre abrem atualizadas.
        if chave == "inicio":
            self._atualizar_inicio()
        elif chave == "estatisticas":
            self._atualizar_estatisticas()
        elif chave == "relatorio":
            self._atualizar_relatorio()

    def _ao_mudar_simulacao(self):
        """Mostra/esconde a faixa amarela e avisa no painel."""
        if self.var_simulacao.get():
            self.faixa_simulacao.pack(side="top", fill="x", before=self._area_paginas)
            log_aviso("Modo simulação LIGADO: as automações vão parar antes de confirmar - nada será gerado.")
        else:
            self.faixa_simulacao.pack_forget()
            log_info("Modo simulação desligado: as automações voltam a confirmar de verdade no SIGEF.")

    def _desligar_simulacao(self):
        if self.automacao_em_andamento:
            messagebox.showinfo("Amigo", "Espere a automação terminar para desligar a simulação.")
            return
        self.var_simulacao.set(False)
        self._ao_mudar_simulacao()

    def _atualizar_status_lateral(self):
        """Rodapé do menu: planilha conectada? automação rodando?"""
        if not hasattr(self, "rotulo_lateral_planilha"):
            return
        if self.dados is not None and self.worksheet is not None:
            self.rotulo_lateral_planilha.config(text=f"📄  Planilha conectada\n      {len(self.dados)} linha(s)")
            _estilo(self.rotulo_lateral_planilha, "LateralStatusOk.TLabel")
        else:
            self.rotulo_lateral_planilha.config(text="📄  Planilha não conectada")
            _estilo(self.rotulo_lateral_planilha, "LateralStatus.TLabel")
        # Não dá para ligar/desligar a simulação no meio de uma automação.
        self.caixa_simulacao.config(state="disabled" if self.automacao_em_andamento else "normal")
        if self.automacao_em_andamento:
            self.rotulo_lateral_robo.config(text="⏳  Trabalhando no SIGEF...\n      não mexa no navegador")
            _estilo(self.rotulo_lateral_robo, "LateralStatusAviso.TLabel")
        else:
            self.rotulo_lateral_robo.config(text="💤  Nenhuma automação rodando")
            _estilo(self.rotulo_lateral_robo, "LateralStatus.TLabel")

    # ---- Gaveta de mensagens (rodapé, recolhível) ----
    def _montar_gaveta_mensagens(self, pai):
        """
        As mensagens do sistema ficam numa "gaveta" no rodapé: fechada, ela
        ocupa 1 linha só e mostra a ÚLTIMA mensagem (colorida); aberta,
        mostra o histórico completo. Antes o quadro de mensagens ficava
        sempre aberto e tomava boa parte da altura da janela.
        """
        gaveta = ttk.Frame(pai, style="Gaveta.TFrame")
        gaveta.pack(side="bottom", fill="x")
        barra = ttk.Frame(gaveta, style="Gaveta.TFrame", padding=px((28, 8, 20, 8)))
        barra.pack(fill="x")
        ttk.Label(barra, text="💬", style="Gaveta.TLabel").pack(side="left")
        self.rotulo_ultima_mensagem = ttk.Label(
            barra, text="As mensagens do programa aparecem aqui.", style="Gaveta.TLabel",
        )
        self.rotulo_ultima_mensagem.pack(side="left", padx=(10, 0), fill="x", expand=True)
        self.botao_gaveta = ttk.Button(barra, text="▴  Ver todas as mensagens", style="Gaveta.TButton",
                                       command=self._alternar_gaveta)
        self.botao_gaveta.pack(side="right")
        Dica(self.botao_gaveta, "Abre (ou fecha) o histórico completo do que o programa fez: "
                                "✅ deu certo, ⚠ atenção, ❗ deu erro.")

        self.corpo_gaveta = ttk.Frame(gaveta, style="Gaveta.TFrame", padding=px((28, 0, 20, 14)))
        barra_rolagem = ttk.Scrollbar(self.corpo_gaveta)
        barra_rolagem.pack(side="right", fill="y")
        self.texto_log = tk.Text(
            self.corpo_gaveta, height=11, wrap="word", state="disabled",
            yscrollcommand=barra_rolagem.set,
        )
        _estilizar_caixa_texto(self.texto_log)
        for nivel, cor in CORES_NIVEL_LOG.items():
            self.texto_log.tag_config(nivel, foreground=cor)
        self.texto_log.pack(fill="both", expand=True)
        barra_rolagem.config(command=self.texto_log.yview)

    def _alternar_gaveta(self):
        self.gaveta_aberta = not self.gaveta_aberta
        if self.gaveta_aberta:
            self.corpo_gaveta.pack(fill="x")
            self.botao_gaveta.config(text="▾  Esconder mensagens")
            self.texto_log.see("end")
            self.mensagens_nao_vistas = 0
        else:
            self.corpo_gaveta.pack_forget()
            self.botao_gaveta.config(text="▴  Ver todas as mensagens")

    def _ao_fechar(self):
        if self.automacao_em_andamento:
            if not messagebox.askyesno(
                "Amigo",
                "Uma automação ainda está trabalhando no SIGEF.\n\n"
                "Se fechar agora, um lançamento pode ficar pela metade.\n\n"
                "Fechar o programa mesmo assim?",
                icon="warning",
            ):
                return
        elif self.itens_manual_resultados:
            if not messagebox.askyesno(
                "Amigo",
                "Os resultados da tela \"✍ Sem planilha\" NÃO ficam salvos em lugar nenhum.\n\n"
                "Se você ainda não anotou ou copiou os números (CE, NL, PP e OB), "
                "clique em \"Não\" e use o botão \"Copiar resultados\".\n\n"
                "Fechar o programa mesmo assim?",
                icon="warning",
            ):
                return
        remover_ouvinte(self._receber_log)
        self.root.destroy()

    # ---- log: thread-safe (automações rodam em background thread) ----
    def _receber_log(self, nivel: str, mensagem: str):
        self.fila_log.put(("log", nivel, mensagem))

    def _drenar_fila_log(self):
        try:
            while True:
                item = self.fila_log.get_nowait()
                if item[0] == "preparo":
                    self._escrever_preparo(item[1])
                elif item[0] == "preparo_fim":
                    self._finalizar_preparo(item[1])
                elif item[0] == "log":
                    self._escrever_log(item[1], item[2])
                elif item[0] == "progresso":
                    self._mostrar_progresso_automacao(item[1], item[2])
        except queue.Empty:
            pass
        self.root.after(100, self._drenar_fila_log)

    def _escrever_preparo(self, mensagem: str):
        if not hasattr(self, "texto_preparo") or not self.texto_preparo.winfo_exists():
            return
        self.texto_preparo.config(state="normal")
        self.texto_preparo.insert("end", mensagem + "\n")
        self.texto_preparo.see("end")
        self.texto_preparo.config(state="disabled")

    def _escrever_log(self, nivel: str, mensagem: str):
        if not hasattr(self, "texto_log") or not self.texto_log.winfo_exists():
            return
        prefixo = PREFIXOS_NIVEL_LOG.get(nivel, "")
        texto = f"{prefixo} {mensagem}" if prefixo else mensagem
        hora = datetime.now().strftime("%H:%M")
        self.texto_log.config(state="normal")
        self.texto_log.insert("end", f"{hora}  {texto}\n", nivel)
        self.texto_log.see("end")
        self.texto_log.config(state="disabled")

        # Gaveta fechada: mostra a última mensagem (cortada, numa linha só).
        resumo = texto if len(texto) <= 150 else texto[:147] + "..."
        if not self.gaveta_aberta:
            self.mensagens_nao_vistas += 1
        self.rotulo_ultima_mensagem.config(text=f"{hora}   {resumo}")
        _estilo(self.rotulo_ultima_mensagem, f"Gaveta{nivel.capitalize()}.TLabel"
                if nivel in CORES_NIVEL_LOG else "Gaveta.TLabel")

    # ------------------------------------------------------------------
    # TELA: INÍCIO
    #
    # Primeira tela ao abrir: em vez de despejar todas as opções de uma vez,
    # pergunta o que a pessoa quer fazer e oferece os 2 caminhos principais
    # em cartões grandes (com planilha / sem planilha), mais um resumo
    # rápido da situação.
    # ------------------------------------------------------------------
    def _montar_aba_inicio(self):
        conteudo = _criar_area_rolavel(self.aba_inicio)
        _criar_cabecalho_aba(
            conteudo,
            "👋  Olá! O que você quer fazer hoje?",
            "Escolha um dos caminhos abaixo. Você pode voltar para cá quando quiser, pelo "
            "🏠 Início no menu da esquerda.",
        )

        caminhos = ttk.Frame(conteudo)
        caminhos.pack(fill="x", pady=(0, 22))
        opcoes = [
            ("📄", "Lançar usando a planilha",
             "O programa lê as pessoas da planilha do Excel e faz as etapas no SIGEF, anotando "
             "os números de volta na planilha.",
             "Melhor para muitas pessoas de uma vez.", "Ir para  🤖 Com planilha  →", "automacoes"),
            ("✍", "Digitar os dados de cada pessoa",
             "Você digita CPF, empenho, banco, agência, conta e valor, e o programa faz sozinho "
             "CE → NL → PP → OB.",
             "Melhor para poucas pessoas. Não usa a planilha.", "Ir para  ✍ Sem planilha  →", "manual"),
        ]
        for coluna, (icone, titulo, texto, melhor, botao_texto, destino) in enumerate(opcoes):
            cartao = Cartao(caminhos)
            cartao.grid(row=0, column=coluna, sticky="nsew", padx=(0 if coluna == 0 else 20, 0))
            caminhos.columnconfigure(coluna, weight=1, uniform="caminhos")
            ttk.Label(cartao, text=icone, style="NumeroGrande.TLabel").pack(anchor="w")
            ttk.Label(cartao, text=titulo, style="Secao.TLabel", wraplength=420, justify="left").pack(anchor="w", pady=(10, 6))
            ttk.Label(cartao, text=texto, style="Instrucao.TLabel", wraplength=420, justify="left").pack(anchor="w")
            ttk.Label(cartao, text=f"👍  {melhor}", style="Dica.TLabel").pack(anchor="w", pady=(12, 18))
            ttk.Button(
                cartao, text=botao_texto, style="Principal.TButton",
                command=lambda d=destino: self._mostrar_pagina(d),
            ).pack(anchor="w")

        resumo = ttk.Frame(conteudo)
        resumo.pack(fill="x", pady=(0, 22))
        self.rotulos_inicio = {}
        blocos = [
            ("parametros", "⚙  Parâmetros", "Conferir parâmetros", "parametros"),
            ("economia", "🎉  Economia", "Ver estatísticas", "estatisticas"),
            ("mes", "📅  Este mês", "Ver relatório", "relatorio"),
        ]
        for coluna, (chave, titulo, botao_texto, destino) in enumerate(blocos):
            cartao = Cartao(resumo, text=titulo)
            cartao.grid(row=0, column=coluna, sticky="nsew", padx=(0 if coluna == 0 else 20, 0))
            resumo.columnconfigure(coluna, weight=1, uniform="resumo")
            valor = ttk.Label(cartao, text="-", style="Subtitulo.TLabel", wraplength=280, justify="left")
            valor.pack(anchor="w")
            detalhe = ttk.Label(cartao, text="", style="Dica.TLabel", wraplength=280, justify="left")
            detalhe.pack(anchor="w", pady=(6, 16))
            ttk.Button(cartao, text=botao_texto, command=lambda d=destino: self._mostrar_pagina(d)).pack(anchor="w")
            self.rotulos_inicio[chave] = (valor, detalhe)

        dica = Cartao(conteudo)
        dica.pack(fill="x")
        ttk.Label(
            dica,
            text="💡  Dica: pare o mouse em cima de qualquer botão para ver o que ele faz. Em cada "
                 "tela, o botão \"❓ Como usar\" (no canto de cima) mostra o passo a passo.",
            style="Instrucao.TLabel", wraplength=900, justify="left",
        ).pack(anchor="w")

    def _atualizar_inicio(self):
        if not getattr(self, "rotulos_inicio", None):
            return
        valor, detalhe = self.rotulos_inicio["parametros"]
        faltando = [nome for chave, nome in (("data", "data"), ("processo", "processo"),
                                             ("mes_referencia", "mês"), ("caminho_planilha", "planilha"))
                    if not str(self.config.get(chave) or "").strip()]
        if faltando:
            valor.config(text=f"Falta preencher: {', '.join(faltando)}")
            _estilo(valor, "Aviso.TLabel")
        else:
            valor.config(text="Tudo preenchido ✅")
            _estilo(valor, "Sucesso.TLabel")
        detalhe.config(
            text=f"Data {self.config.get('data') or '-'}  •  Mês {self.config.get('mes_referencia') or '-'}\n"
                 f"Processo {self.config.get('processo') or '-'}"
        )

        valor, detalhe = self.rotulos_inicio["economia"]
        try:
            totais = calcular_estatisticas("Todo o período", self.config)["totais"]
            valor.config(text=formatar_duracao(max(0, totais["tempo_economizado_segundos"])))
            detalhe.config(text=f"desde o começo, em {totais['processados']} item(ns) feitos pela automação")
        except Exception:
            valor.config(text="-")
        _estilo(valor, "Sucesso.TLabel")

        valor, detalhe = self.rotulos_inicio["mes"]
        try:
            mes = resumo_mes(datetime.now().strftime("%m/%Y"))
            partes = [f"{categoria} {mes['categorias'][categoria]['processados']}" for categoria in CATEGORIAS_PRINCIPAIS]
            valor.config(text="   •   ".join(partes))
            detalhe.config(text=f"{mes['total_execucoes']} automação(ões) rodada(s) em {mes['mes_ano']}")
        except Exception:
            valor.config(text="-")

    # ------------------------------------------------------------------
    # ABA: AUTOMAÇÕES
    # ------------------------------------------------------------------
    def _montar_aba_automacoes(self):
        # A aba fica dentro de uma área com rolagem (ver `_criar_area_rolavel`)
        # - com fontes maiores e o painel de andamento, o conteúdo desta aba
        # passou a ser mais alto do que cabe de uma vez em telas menores;
        # assim, nada fica escondido/cortado embaixo - basta rolar para ver
        # o resto.
        conteudo = _criar_area_rolavel(self.aba_automacoes)

        _criar_cabecalho_aba(
            conteudo,
            "🤖  Lançar com a planilha",
            "Conecte a planilha do Excel e clique nas etapas na ordem, de cima para baixo. "
            "Os números gerados vão direto para a planilha.",
            passos=[
                "Confira na tela ⚙ Parâmetros a data, o processo, o mês e qual é a planilha.",
                "Deixe o navegador do SIGEF aberto e com o login feito.",
                "Clique em \"🔌 Conectar planilha\" e espere aparecer ✅ em verde.",
                "Clique nas etapas uma de cada vez, nesta ordem: CE → NL → PP → Conferir contas → OB. "
                "Espere cada uma terminar antes de clicar na próxima.",
                "Acompanhe o quadro \"Andamento\" e as mensagens lá embaixo. Passe o mouse em cima "
                "de um botão para ver o que ele faz.",
            ],
            ajuda_aberta=False,
        )

        # ---- Planilha ----
        # Planilha e Andamento lado a lado: as 2 coisas que a pessoa olha o
        # tempo todo ficam no topo, sem empurrar as etapas para baixo.
        linha_topo = ttk.Frame(conteudo)
        linha_topo.pack(fill="x", pady=(0, 22))
        linha_topo.columnconfigure(0, weight=1, uniform="topo")
        linha_topo.columnconfigure(1, weight=1, uniform="topo")
        frame_planilha = Cartao(linha_topo, text="📄  Planilha")
        frame_planilha.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.rotulo_status_planilha = ttk.Label(
            frame_planilha, text="⚪  Nenhuma planilha conectada ainda.", style="Instrucao.TLabel",
            wraplength=420, justify="left",
        )
        self.rotulo_status_planilha.pack(anchor="w", pady=(0, 14))
        botao_conectar = ttk.Button(
            frame_planilha, text="🔌  Conectar planilha", style="Principal.TButton",
            command=self._conectar_e_ler,
        )
        botao_conectar.pack(anchor="w")
        Dica(botao_conectar, "Abre (ou reaproveita, se já estiver aberta) a planilha escolhida na tela "
                             "⚙ Parâmetros e lê as linhas a partir da \"Linha inicial\".")
        self.botao_conectar = botao_conectar

        # ---- Painel de andamento (barra de progresso das automações) ----
        # A barra fica "indeterminada" (vai e volta) em vez de mostrar uma
        # porcentagem exata, porque as automações não informam quantas
        # etapas internas faltam - ainda assim, deixa bem claro quando algo
        # está rodando e quando terminou (com sucesso ou erro).
        frame_progresso = Cartao(linha_topo, text="⏱  Andamento")
        frame_progresso.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.rotulo_progresso = ttk.Label(
            frame_progresso, text="💤  Nenhuma automação rodando agora.",
            style="Instrucao.TLabel", wraplength=420, justify="left",
        )
        self.rotulo_progresso.pack(anchor="w", pady=(0, 14))
        # Parada, a barra fica vazia (modo "determinado", valor 0); rodando,
        # vira a barra que vai e volta (ver `_iniciar_progresso_automacao`).
        self.barra_progresso_automacao = ttk.Progressbar(frame_progresso, mode="determinate", value=0)
        self.barra_progresso_automacao.pack(fill="x", pady=(8, 0))
        self.botao_parar_automacao = ttk.Button(
            frame_progresso, text="⏹  Parar depois da linha atual", state="disabled",
            command=self._parar_automacao_planilha,
        )
        self.botao_parar_automacao.pack(anchor="w", pady=(14, 0))
        Dica(self.botao_parar_automacao,
             "A linha que está sendo feita agora termina normalmente (nada fica pela metade) e a "
             "automação para antes da próxima. Na OB, para entre um lote e outro.")

        for indice, (titulo, numeros) in enumerate(GRUPOS_AUTOMACOES):
            grupo = Cartao(conteudo, text=f"{NUMEROS_CIRCULADOS[indice]}  {titulo}", padding=12)
            grupo.pack(fill="x", pady=(0, 18))
            ttk.Label(grupo, text=DESCRICOES_GRUPOS_AUTOMACOES[titulo], style="Dica.TLabel").pack(
                anchor="w", pady=(0, 8)
            )
            linha_botoes = ttk.Frame(grupo)
            linha_botoes.pack(fill="x")
            for numero in numeros:
                texto_botao, explicacao = BOTOES_AUTOMACOES[numero]
                botao = ttk.Button(
                    linha_botoes, text=texto_botao, state="disabled",
                    command=lambda n=numero: self._executar_automacao(n),
                )
                botao.pack(side="left", padx=(0, 10), pady=2)
                Dica(botao, explicacao)
                self.botoes_automacao[numero] = botao

    def _definir_botoes_automacao(self, habilitado: bool):
        estado = "normal" if habilitado else "disabled"
        for botao in self.botoes_automacao.values():
            botao.config(state=estado)

    def _parar_automacao_planilha(self):
        execucao.pedir_parada()
        self.botao_parar_automacao.config(state="disabled", text="⏳  Vai parar depois da linha atual...")
        log_aviso("Parada pedida: a automação vai terminar a linha atual e parar.")

    def _mostrar_progresso_automacao(self, atual: int, total: int):
        """Recebe o "linha X de Y" das automações da planilha (o modo Sem
        planilha tem o próprio andamento, por pessoa)."""
        if not self.automacao_em_andamento or self._manual_rodando or total <= 0:
            return
        barra = self.barra_progresso_automacao
        if str(barra.cget("mode")) != "determinate":
            barra.stop()
            barra.config(mode="determinate")
        barra.config(maximum=total, value=max(0, atual - 1))
        prefixo = "🧪  Simulando" if execucao.em_simulacao() else "⏳  Fazendo"
        texto = f"{prefixo} \"{self._nome_automacao_atual}\": linha {min(atual, total)} de {total}."
        if atual >= total:
            texto = f"{prefixo} \"{self._nome_automacao_atual}\": terminando..."
        _configurar(self.rotulo_progresso, text=texto + " Não mexa no navegador.", style="Subtitulo.TLabel")

    def _iniciar_progresso_automacao(self, nome: str):
        _configurar(self.rotulo_progresso, 
            text=f"⏳  Fazendo \"{nome}\"... isso pode levar alguns minutos. Aguarde e não mexa no navegador.",
            style="Subtitulo.TLabel",
        )
        self.barra_progresso_automacao.config(mode="indeterminate")
        self.barra_progresso_automacao.start(12)

    def _parar_progresso_automacao(self, sucesso: bool, parou: bool = False, simulacao: bool = False):
        self.barra_progresso_automacao.stop()
        self.barra_progresso_automacao.config(mode="determinate", value=0)
        self.botao_parar_automacao.config(state="disabled", text="⏹  Parar depois da linha atual")
        if sucesso and parou:
            _configurar(self.rotulo_progresso,
                        text="⏹  Parado a seu pedido. As linhas que faltaram não foram feitas.",
                        style="Aviso.TLabel")
            return
        if sucesso and simulacao:
            _configurar(self.rotulo_progresso,
                        text="🧪  Simulação terminada - nada foi gerado. Veja na planilha o que daria certo (✓) e errado (✗).",
                        style="Aviso.TLabel")
            return
        if sucesso:
            _configurar(self.rotulo_progresso, 
                text="✅  Pronto! A automação terminou. Pode clicar na próxima etapa.",
                style="Sucesso.TLabel",
            )
        else:
            _configurar(self.rotulo_progresso, 
                text="❗  A automação terminou com erro - veja nas mensagens lá embaixo o que aconteceu.",
                style="Erro.TLabel",
            )

    def _conectar_e_ler(self):
        self.botao_conectar.config(state="disabled")
        self._definir_botoes_automacao(False)
        threading.Thread(target=self._conectar_e_ler_em_segundo_plano, daemon=True).start()

    def _conectar_e_ler_em_segundo_plano(self):
        try:
            worksheet = conectar_planilha(self.config)
            if worksheet is None:
                self.root.after(0, lambda: self.botao_conectar.config(state="normal"))
                return

            linha_inicial = self.config.get("linha_inicial", 2)
            dados = ler_dados(worksheet, linha_inicial)

            self.worksheet = worksheet
            self.dados = dados

            if not dados:
                log_erro("Nenhum dado encontrado para processar (verifique a linha inicial).")
                self.root.after(0, lambda: self.botao_conectar.config(state="normal"))
                return

            def atualizar():
                _configurar(self.rotulo_status_planilha, 
                    text=f"✅  Planilha conectada: {len(dados)} linha(s) lida(s), a partir da linha {linha_inicial}.",
                    style="Sucesso.TLabel",
                )
                self.botao_conectar.config(state="normal")
                # Se o Lançamento Manual estiver rodando, os botões só
                # voltam quando ele terminar (1 automação por vez no SIGEF).
                self._definir_botoes_automacao(not self.automacao_em_andamento)
                self._atualizar_status_lateral()

            self.root.after(0, atualizar)
        except Exception:
            log_erro("Erro inesperado ao conectar/ler a planilha:")
            for linha in traceback.format_exc().splitlines():
                log_erro(linha)
            self.root.after(0, lambda: self.botao_conectar.config(state="normal"))

    def _executar_automacao(self, numero: int):
        nome, funcao = PROGRAMAS[numero]
        if self.dados is None or self.worksheet is None:
            messagebox.showinfo("Amigo", "Conecte a planilha e leia os dados primeiro.")
            return
        if self.automacao_em_andamento:
            messagebox.showinfo("Amigo", "Espere a automação que já está rodando terminar.")
            return

        simulacao = self.var_simulacao.get()
        if simulacao:
            pergunta = (
                f"🧪 SIMULAR a automação '{nome}' agora?\n\n"
                "Ela vai preencher e conferir tudo no SIGEF de verdade, mas vai PARAR ANTES de "
                "confirmar - nada será gerado.\n\nNa planilha, no lugar do número, vai aparecer "
                "\"SIMULADO ✓\" (daria certo) ou \"SIMULADO ✗\" com o motivo."
            )
        else:
            pergunta = (
                f"Executar a automação '{nome}' agora?\n\nIsso vai interagir com o navegador e "
                "com a planilha em tempo real, gerando documentos DE VERDADE no SIGEF."
            )
        if not messagebox.askyesno("Amigo", pergunta):
            return

        execucao.iniciar_execucao(simulacao=simulacao)
        self._nome_automacao_atual = nome
        self.automacao_em_andamento = True
        self._atualizar_botoes_manual()
        self._atualizar_status_lateral()
        self._definir_botoes_automacao(False)
        self.botao_conectar.config(state="disabled")
        log_info(f"Iniciando {'SIMULAÇÃO da ' if simulacao else ''}automação: {nome}...")
        self._iniciar_progresso_automacao(nome)
        self.botao_parar_automacao.config(state="normal", text="⏹  Parar depois da linha atual")

        def rodar():
            sucesso = False
            processados = 0
            detalhe = None
            inicio = datetime.now()
            try:
                resultado = funcao(self.dados, self.config, self.worksheet)
                # A maioria das automações devolve só a quantidade
                # processada (int); "Gerar OB" devolve uma tupla
                # (itens confirmados, nº de OB's geradas), já que 1 OB
                # agrupa várias linhas da planilha - ver amigo/automacoes/ob.py.
                if isinstance(resultado, tuple):
                    processados, total_obs = resultado
                    detalhe = f"{total_obs} OB(s) gerada(s)"
                else:
                    processados = resultado or 0
                log_sucesso(f"Automação '{nome}' finalizada.")
                sucesso = True
            except Exception:
                log_erro(f"Erro inesperado durante a automação '{nome}':")
                for linha in traceback.format_exc().splitlines():
                    log_erro(linha)
            finally:
                fim = datetime.now()
                parou = execucao.parada_pedida()
                execucao.finalizar_execucao()
                # O registro do histórico NUNCA pode derrubar a automação
                # que acabou de rodar (ela já terminou nesse ponto, mas o
                # padrão do resto do sistema é sempre blindar contra
                # imprevisto de E/S - disco cheio, permissão etc.).
                # Simulação não gera nada, então não entra no Relatório nem
                # nas Estatísticas (senão distorceria o tempo economizado).
                if simulacao:
                    log_info("Simulação não entra no Relatório nem nas Estatísticas.")
                else:
                    try:
                        registrar_execucao(
                            programa=nome, inicio=inicio, fim=fim,
                            processados=processados, total_linhas=len(self.dados or []),
                            sucesso=sucesso, detalhe=detalhe,
                        )
                    except Exception:
                        log_erro("Não foi possível salvar esta execução no histórico do Relatório.")

                def reabilitar():
                    self.automacao_em_andamento = False
                    self._atualizar_botoes_manual()
                    self._atualizar_status_lateral()
                    self._definir_botoes_automacao(True)
                    self.botao_conectar.config(state="normal")
                    self._parar_progresso_automacao(sucesso, parou=parou, simulacao=simulacao)
                    self._atualizar_relatorio()
                    self._atualizar_estatisticas()
                self.root.after(0, reabilitar)

        threading.Thread(target=rodar, daemon=True).start()

    # ------------------------------------------------------------------
    # ABA: LANÇAMENTO MANUAL (sem planilha)
    #
    # Alternativa à planilha: o usuário digita os dados de cada pessoa e
    # o programa faz CE -> NL -> PP -> OB, uma pessoa de cada vez (a CE
    # gerada vai para a NL, a NL para a PP e a PP para a OB). Toda a lógica
    # está em `amigo.lancamento_manual`, que roda as MESMAS automações da
    # planilha usando uma "planilha" só em memória - nada é gravado no
    # Excel, em arquivo ou no histórico. Esta aba não usa nem altera
    # `self.dados`/`self.worksheet` (os da planilha conectada).
    # ------------------------------------------------------------------
    def _montar_campo(self, pai, linha: int, coluna: int, titulo: str, dica: str, var, largura: int = 22):
        """
        1 campo de formulário no padrão de todas as telas: rótulo EM CIMA,
        caixa de digitar no meio e, embaixo, o texto de ajuda (exemplo, ou
        a conferência do que foi digitado). Devolve (caixa, texto_de_ajuda).
        """
        bloco = ttk.Frame(pai)
        bloco.grid(row=linha, column=coluna, sticky="nwe", padx=(0 if coluna == 0 else 28, 0), pady=8)
        # Colunas do formulário dividem a largura por igual: caixas mais
        # largas e alinhadas, sem sobra vazia à direita.
        pai.columnconfigure(coluna, weight=1, uniform="campos")
        ttk.Label(bloco, text=titulo, style="Campo.TLabel").pack(anchor="w")
        entrada = ttk.Entry(bloco, textvariable=var, width=largura)
        entrada.pack(fill="x", pady=(6, 4))
        retorno = _texto_ajuda_campo(bloco, "")
        retorno.config(wraplength=300)
        retorno.pack(anchor="w")
        _mostrar_retorno_campo(retorno, "dica", dica)
        return entrada, retorno

    def _montar_observacoes(self, pai, valores_salvos: dict, obter_mes_processo) -> dict:
        """
        Bloco "Textos de observação" (CE e OB), usado nos Parâmetros e no
        modo Sem planilha: para cada texto, a pessoa escolhe entre o texto
        padrão (o de sempre) e um texto próprio, e vê embaixo exatamente
        como ele vai ficar no SIGEF. `obter_mes_processo()` devolve
        (mês referência, processo) digitados na mesma tela, para a prévia.
        """
        controles = {}
        ttk.Label(
            pai,
            text="Escreva {mes} onde deve ir o mês referência e {processo} onde deve ir o número "
                 "do processo - o programa troca sozinho, todo mês.",
            style="Dica.TLabel", wraplength=900, justify="left",
        ).pack(anchor="w", pady=(0, 8))

        for tipo in TIPOS_OBSERVACAO:
            bloco = ttk.Frame(pai)
            bloco.pack(fill="x", pady=(8, 6))
            ttk.Label(bloco, text=f"📝  {NOMES_OBSERVACAO[tipo]}", style="Campo.TLabel").pack(anchor="w")
            personalizado = str(valores_salvos.get(CHAVES_CONFIG_OBSERVACAO[tipo]) or "").strip()
            modo = tk.StringVar(value="proprio" if personalizado else "padrao")
            texto = tk.StringVar(value=personalizado or TEXTO_PADRAO_OBSERVACAO[tipo])

            opcoes = ttk.Frame(bloco)
            opcoes.pack(anchor="w", pady=(4, 2))
            entrada = ttk.Entry(bloco, textvariable=texto)
            previa = _texto_ajuda_campo(bloco, "")
            previa.config(wraplength=900)

            def atualizar(*_args, tipo=tipo):
                self._atualizar_observacao(controles[tipo], tipo, obter_mes_processo)

            ttk.Radiobutton(opcoes, text="Usar o texto padrão (recomendado)", value="padrao",
                            variable=modo, command=atualizar).pack(side="left")
            ttk.Radiobutton(opcoes, text="Escrever um texto próprio", value="proprio",
                            variable=modo, command=atualizar).pack(side="left", padx=(28, 0))
            entrada.pack(fill="x", pady=(6, 4))
            previa.pack(anchor="w")
            texto.trace_add("write", atualizar)
            controles[tipo] = {"modo": modo, "texto": texto, "entrada": entrada, "previa": previa}
            atualizar()
        return controles

    def _atualizar_observacao(self, controle: dict, tipo: str, obter_mes_processo) -> None:
        """Liga/desliga a caixa de texto e mostra a prévia (✅/⚠/❗)."""
        padrao = controle["modo"].get() == "padrao"
        if padrao:
            if controle["texto"].get() != TEXTO_PADRAO_OBSERVACAO[tipo]:
                controle["texto"].set(TEXTO_PADRAO_OBSERVACAO[tipo])  # (dispara de novo esta função)
                return
            controle["entrada"].config(state="disabled")
        else:
            controle["entrada"].config(state="normal")
        modelo = controle["texto"].get().strip()
        if not modelo:
            _mostrar_retorno_campo(controle["previa"], "erro", "Escreva o texto, ou marque \"Usar o texto padrão\".")
            return
        mes, processo = obter_mes_processo()
        final = aplicar_marcacoes(modelo, mes or "{mes}", processo or "{processo}")
        if len(final) > LIMITE_AVISO_CARACTERES:
            _mostrar_retorno_campo(
                controle["previa"], "aviso",
                f"Texto longo ({len(final)} letras) - o SIGEF pode cortar. Vai ficar assim: {final}",
            )
        else:
            _mostrar_retorno_campo(controle["previa"], "ok", f"Vai ficar assim no SIGEF: {final}")

    @staticmethod
    def _valores_observacoes(controles: dict):
        """Devolve ({chave_config: texto}, [erros]). Texto padrão = ""
        (vazio), que é como a configuração guarda "usar o padrão"."""
        valores, erros = {}, []
        for tipo, controle in controles.items():
            chave = CHAVES_CONFIG_OBSERVACAO[tipo]
            if controle["modo"].get() == "padrao":
                valores[chave] = ""
                continue
            texto = controle["texto"].get().strip()
            if not texto:
                erros.append(f"- {NOMES_OBSERVACAO[tipo]}: escreva o texto ou marque \"Usar o texto padrão\".")
            valores[chave] = "" if texto == TEXTO_PADRAO_OBSERVACAO[tipo] else texto
        return valores, erros

    def _montar_aba_manual(self):
        conteudo = _criar_area_rolavel(self.aba_manual)

        _criar_cabecalho_aba(
            conteudo,
            "✍  Lançar sem planilha",
            "Digite os dados de cada pessoa e o programa faz sozinho CE → NL → PP → OB. "
            "Não usa e não mexe na planilha do Excel.",
            passos=[
                "Preencha os parâmetros deste lançamento (data, processo e mês). Eles valem só "
                "para esta aba.",
                "Deixe o navegador do SIGEF aberto e com o login feito.",
                "Digite os dados da pessoa e clique em \"➕ Adicionar à lista\". Repita para "
                "cada pessoa.",
                "Clique em \"▶ Começar lançamentos\". O programa faz uma pessoa de cada vez: "
                "CE, depois NL, depois PP e por último OB.",
                "Os números aparecem em \"Resultados\". Eles NÃO ficam salvos: clique em "
                "\"📋 Copiar resultados\" ou anote antes de fechar o programa.",
            ],
            ajuda_aberta=False,
        )

        # ---- ① Parâmetros próprios deste lançamento ----
        # Os mesmos 3 campos da aba "Parâmetros", no mesmo formato - mas
        # guardados à parte (config.json, chave
        # `CHAVE_CONFIG_PARAMETROS_MANUAL`), sem tocar nos da planilha.
        frame_parametros = Cartao(conteudo, text="①  Parâmetros deste lançamento", padding=14)
        frame_parametros.pack(fill="x", pady=(0, 22))
        ttk.Label(
            frame_parametros,
            text="Valem para todas as pessoas da lista e ficam guardados para a próxima vez. "
                 "São separados dos da tela ⚙ Parâmetros.",
            style="Dica.TLabel", wraplength=980, justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        salvos = parametros_manuais_salvos(self.config)
        for indice, (chave, titulo, dica) in enumerate(CAMPOS_PARAMETROS_MANUAL):
            var = tk.StringVar(value=salvos.get(chave, ""))
            entrada, retorno = self._montar_campo(
                frame_parametros, 1, indice, titulo, dica, var, largura=26 if chave == "processo" else 18,
            )
            var.trace_add("write", lambda *_args: self._conferir_parametros_manual())
            self.vars_parametros_manual[chave] = var
            self.entradas_parametros_manual[chave] = entrada
            self.retornos_parametros_manual[chave] = (retorno, dica)

        # Textos de observação (padrão ou próprio) - só deste lançamento.
        area_observacoes = ttk.Frame(frame_parametros)
        area_observacoes.grid(row=2, column=0, columnspan=3, sticky="we", pady=(12, 0))
        self.observacoes_manual = self._montar_observacoes(
            area_observacoes, salvos,
            lambda: (self.vars_parametros_manual["mes_referencia"].get().strip(),
                     self.vars_parametros_manual["processo"].get().strip()),
        )

        botoes_parametros = ttk.Frame(frame_parametros)
        botoes_parametros.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
        botao_salvar = ttk.Button(
            botoes_parametros, text="💾  Salvar parâmetros", command=self._salvar_parametros_manual,
        )
        botao_salvar.pack(side="left")
        Dica(botao_salvar, "Guarda a data, o processo e o mês neste computador, para já aparecerem "
                           "preenchidos da próxima vez. (Ao começar os lançamentos eles também são guardados.)")
        botao_copiar = ttk.Button(
            botoes_parametros, text="📥  Usar os mesmos da tela Parâmetros",
            command=self._copiar_parametros_da_planilha_manual,
        )
        botao_copiar.pack(side="left", padx=(10, 0))
        Dica(botao_copiar, "Preenche estes 3 campos com os mesmos valores que estão na tela ⚙ Parâmetros "
                           "(os da planilha). Útil quando são iguais.")
        self._conferir_parametros_manual()

        # ---- ② Formulário: dados de 1 pessoa ----
        frame_formulario = Cartao(conteudo, text="②  Dados da pessoa", padding=14)
        frame_formulario.pack(fill="x", pady=(0, 22))
        ttk.Label(
            frame_formulario,
            text="Embaixo de cada campo aparece um exemplo. Ao digitar, ele mostra ✅ como o dado "
                 "vai ser usado, ou ❗ o que corrigir. A tecla Enter passa para o próximo campo.",
            style="Dica.TLabel", wraplength=980, justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        for indice, (campo, titulo, dica) in enumerate(CAMPOS_TELA_MANUAL):
            var = tk.StringVar()
            entrada, retorno = self._montar_campo(frame_formulario, 1 + indice // 3, indice % 3, titulo, dica, var)
            var.trace_add("write", lambda *_args, c=campo: self._conferir_campo_manual(c))
            entrada.bind("<Return>", lambda _evento, i=indice: self._proximo_campo_manual(i))
            self.vars_manual[campo] = var
            self.entradas_manual[campo] = entrada
            self.retornos_manual[campo] = (retorno, dica)

        botoes_formulario = ttk.Frame(frame_formulario)
        botoes_formulario.grid(row=3, column=0, columnspan=3, sticky="w", pady=(12, 0))
        botao_adicionar = ttk.Button(
            botoes_formulario, text="➕  Adicionar à lista", style="Principal.TButton",
            command=self._adicionar_item_manual,
        )
        botao_adicionar.pack(side="left")
        Dica(botao_adicionar, "Confere os 6 campos e coloca a pessoa na lista abaixo. Nada é lançado "
                              "no SIGEF ainda - isso só acontece em \"▶ Começar lançamentos\".")
        ttk.Button(
            botoes_formulario, text="🧹  Limpar campos", command=self._limpar_formulario_manual,
        ).pack(side="left", padx=(10, 0))

        # ---- ③ Lista de pessoas para lançar ----
        frame_lista = Cartao(conteudo, text="③  Lista de pessoas para lançar", padding=14)
        frame_lista.pack(fill="x", pady=(0, 22))
        self.rotulo_lista_manual = ttk.Label(frame_lista, text="", style="Dica.TLabel")
        self.rotulo_lista_manual.pack(anchor="w", pady=(0, 8))

        self.tabela_pendentes_manual = self._criar_tabela_manual(frame_lista, [
            ("numero", "Nº", 45), ("cpf", "CPF", 150), ("ne", "Nota de Empenho", 195),
            ("banco", "Banco", 80), ("agencia", "Agência", 100), ("conta", "Conta", 120),
            ("valor", "Valor", 140), ("observacao", "Observação", 240),
        ])

        botoes_lista = ttk.Frame(frame_lista)
        botoes_lista.pack(fill="x", pady=(10, 0))
        self.botao_corrigir_manual = ttk.Button(
            botoes_lista, text="✏  Corrigir", command=self._corrigir_item_manual,
        )
        self.botao_corrigir_manual.pack(side="left")
        Dica(self.botao_corrigir_manual, "Clique primeiro na pessoa da tabela. Os dados dela voltam para os "
                                         "campos acima para você corrigir e adicionar de novo.")
        self.botao_remover_manual = ttk.Button(
            botoes_lista, text="🗑  Tirar da lista", command=self._remover_item_manual,
        )
        self.botao_remover_manual.pack(side="left", padx=(10, 0))
        Dica(self.botao_remover_manual, "Clique primeiro na pessoa da tabela. Ela sai da lista e não é lançada.")
        self.botao_esvaziar_manual = ttk.Button(
            botoes_lista, text="🧺  Esvaziar a lista", command=self._esvaziar_lista_manual,
        )
        self.botao_esvaziar_manual.pack(side="left", padx=(10, 0))
        Dica(self.botao_esvaziar_manual, "Tira TODAS as pessoas da lista (pede confirmação antes).")

        # ---- ④ Execução ----
        frame_execucao = Cartao(conteudo, text="④  Lançar no SIGEF", padding=14)
        frame_execucao.pack(fill="x", pady=(0, 22))
        botoes_execucao = ttk.Frame(frame_execucao)
        botoes_execucao.pack(fill="x", pady=(0, 10))
        self.botao_comecar_manual = ttk.Button(
            botoes_execucao, text="▶  Começar lançamentos", style="Principal.TButton",
            command=self._comecar_lancamentos_manual,
        )
        self.botao_comecar_manual.pack(side="left")
        Dica(self.botao_comecar_manual, "Lança DE VERDADE no SIGEF todas as pessoas da lista, uma de cada vez "
                                        "(CE → NL → PP → OB). Antes de começar, o programa mostra um resumo e "
                                        "pede confirmação.")
        self.botao_parar_manual = ttk.Button(
            botoes_execucao, text="⏸  Parar depois da pessoa atual", state="disabled",
            command=self._pedir_parada_manual,
        )
        self.botao_parar_manual.pack(side="left", padx=(10, 0))
        Dica(self.botao_parar_manual, "A pessoa que está sendo feita agora termina normalmente. As que ainda "
                                      "não começaram voltam para a lista.")
        self.rotulo_andamento_manual = ttk.Label(
            frame_execucao, text="💤  Nenhum lançamento rodando agora.", style="Instrucao.TLabel",
            wraplength=980, justify="left",
        )
        self.rotulo_andamento_manual.pack(anchor="w", pady=(0, 8))
        self.barra_manual = ttk.Progressbar(frame_execucao, mode="determinate")
        self.barra_manual.pack(fill="x")

        # ---- ⑤ Resultados ----
        frame_resultados = Cartao(conteudo, text="⑤  Resultados", padding=14)
        frame_resultados.pack(fill="x", pady=(0, 22))
        ttk.Label(
            frame_resultados,
            text="⚠  Estes números NÃO ficam salvos - copie ou anote antes de fechar o programa.   "
                 "Clique numa linha para ver os detalhes.",
            style="DicaAviso.TLabel", wraplength=980, justify="left",
        ).pack(anchor="w", pady=(0, 8))
        self.tabela_resultados_manual = self._criar_tabela_manual(frame_resultados, [
            ("numero", "Nº", 50), ("cpf", "CPF", 150), ("valor", "Valor", 130),
            ("ce", "CE", 140), ("nl", "NL", 140), ("pp", "PP", 140), ("ob", "OB", 140),
            ("situacao", "Situação", 170),
        ])
        self.tabela_resultados_manual.bind(
            "<<TreeviewSelect>>", lambda _evento: self._mostrar_detalhe_resultado_manual()
        )
        self.rotulo_detalhe_manual = ttk.Label(
            frame_resultados, text="", style="Instrucao.TLabel", wraplength=980, justify="left",
        )
        self.rotulo_detalhe_manual.pack(anchor="w", pady=(8, 0))

        botoes_resultados = ttk.Frame(frame_resultados)
        botoes_resultados.pack(fill="x", pady=(10, 0))
        self.botao_copiar_manual = ttk.Button(
            botoes_resultados, text="📋  Copiar resultados", style="Principal.TButton",
            command=self._copiar_resultados_manual,
        )
        self.botao_copiar_manual.pack(side="left")
        Dica(self.botao_copiar_manual, "Copia a tabela inteira. Depois é só colar (Ctrl + V) no Excel, no "
                                       "Word ou no bloco de notas.")
        self.botao_repetir_manual = ttk.Button(
            botoes_resultados, text="🔁  Tentar de novo as que pararam",
            command=self._repetir_itens_parados_manual,
        )
        self.botao_repetir_manual.pack(side="left", padx=(10, 0))
        Dica(self.botao_repetir_manual, "As pessoas com ❗ voltam para a lista e continuam da etapa em que "
                                        "pararam - o que já deu certo não é feito de novo (nada em dobro no SIGEF).")
        self.botao_limpar_resultados_manual = ttk.Button(
            botoes_resultados, text="🧽  Apagar da tela",
            command=self._limpar_resultados_manual,
        )
        self.botao_limpar_resultados_manual.pack(side="left", padx=(10, 0))
        Dica(self.botao_limpar_resultados_manual, "Tira os resultados da tela. Como eles não ficam salvos, "
                                                  "copie antes!")

        self._atualizar_lista_manual()

    def _criar_tabela_manual(self, pai, colunas_tabela) -> ttk.Treeview:
        """Tabela (Treeview) com barra de rolagem, no tema escuro. Cores por
        situação: verde (concluído), vermelho (parou), amarelo (andamento)."""
        frame = ttk.Frame(pai)
        frame.pack(fill="x")
        tabela = ttk.Treeview(
            frame, columns=[chave for chave, _t, _l in colunas_tabela],
            show="headings", height=5, selectmode="browse",
        )
        for chave, titulo, largura in colunas_tabela:
            tabela.heading(chave, text=titulo, anchor="w")
            # Colunas curtas (Nº, banco, agência) ficam fixas; as outras
            # dividem o espaço que sobrar.
            fixa = chave in ("numero", "banco", "agencia")
            tabela.column(chave, width=px(largura), minwidth=px(largura if fixa else 60), anchor="w", stretch=not fixa)
        rolagem = ttk.Scrollbar(frame, orient="vertical", command=tabela.yview)
        tabela.configure(yscrollcommand=rolagem.set)
        tabela.pack(side="left", fill="x", expand=True)
        rolagem.pack(side="right", fill="y")
        tabela.tag_configure("concluido", foreground=COR_VERDE)
        tabela.tag_configure("parou", foreground=COR_VERMELHO)
        tabela.tag_configure("andamento", foreground=COR_AMARELO)
        return tabela

    # ---- parâmetros deste lançamento ----
    def _conferir_parametros_manual(self):
        """Mesmo retorno visual dos campos da pessoa: exemplo (vazio),
        verde (como vai ser usado) ou vermelho (o que corrigir). O mês
        referência ainda ganha um lembrete amarelo quando não é o mês
        anterior ao da Data (a mesma regra da aba "Parâmetros")."""
        textos = {chave: var.get() for chave, var in self.vars_parametros_manual.items()}
        for tipo, controle in getattr(self, "observacoes_manual", {}).items():
            self._atualizar_observacao(
                controle, tipo,
                lambda: (textos["mes_referencia"].strip(), textos["processo"].strip()),
            )
        for chave, texto in textos.items():
            rotulo, dica = self.retornos_parametros_manual[chave]
            if not texto.strip():
                _mostrar_retorno_campo(rotulo, "dica", dica)
                continue
            try:
                _valor, exibicao = padronizar_parametro_manual(chave, texto)
            except ValueError as erro:
                _mostrar_retorno_campo(rotulo, "erro", str(erro))
                continue
            aviso = aviso_mes_referencia(textos["data"], texto) if chave == "mes_referencia" else ""
            if aviso:
                _mostrar_retorno_campo(rotulo, "aviso", f"{exibicao}. {aviso}")
            else:
                _mostrar_retorno_campo(rotulo, "ok", f"Certo: {exibicao}")

    def _parametros_manuais_da_tela(self, mostrar_erros: bool = True):
        """Padroniza os 3 parâmetros digitados. Devolve o dicionário pronto
        ou None (mostrando o que corrigir)."""
        textos = {chave: var.get() for chave, var in self.vars_parametros_manual.items()}
        valores_observacoes, erros_observacoes = self._valores_observacoes(self.observacoes_manual)
        textos.update(valores_observacoes)
        parametros, erros = montar_parametros_manuais(textos)
        if erros_observacoes and mostrar_erros:
            messagebox.showwarning(
                "Amigo", "Confira os textos de observação:\n\n" + "\n".join(erros_observacoes),
            )
            return None
        if erros and mostrar_erros:
            linhas = [
                f"- {TITULOS_PARAMETROS_MANUAL[chave]}: {erros[chave]}"
                for chave, _t, _d in CAMPOS_PARAMETROS_MANUAL if chave in erros
            ]
            messagebox.showwarning(
                "Amigo",
                "Confira os \"Parâmetros deste lançamento\" (o que está com ❗):\n\n" + "\n".join(linhas),
            )
            primeiro = next(chave for chave, _t, _d in CAMPOS_PARAMETROS_MANUAL if chave in erros)
            self.entradas_parametros_manual[primeiro].focus_set()
        return parametros

    def _guardar_parametros_manual(self, parametros: dict):
        """Grava em config.json, na chave própria desta aba - os parâmetros
        da planilha ("data", "processo", "mes_referencia") ficam intactos."""
        self.config[CHAVE_CONFIG_PARAMETROS_MANUAL] = dict(parametros)
        salvar_configuracoes(self.config)

    def _salvar_parametros_manual(self):
        parametros = self._parametros_manuais_da_tela()
        if parametros is None:
            return
        self._guardar_parametros_manual(parametros)
        for chave, var in self.vars_parametros_manual.items():
            var.set(parametros[chave])
        log_sucesso("Parâmetros do Lançamento Manual salvos (os da planilha não mudaram).")

    def _copiar_parametros_da_planilha_manual(self):
        if not messagebox.askyesno(
            "Amigo",
            "Trocar a Data, o Processo, o Mês referência e os textos de observação deste quadro "
            "pelos que estão na tela \"⚙ Parâmetros\"?\n\n(Depois, se quiser guardar, clique em \"💾 Salvar parâmetros\".)",
        ):
            return
        for chave, var in self.vars_parametros_manual.items():
            var.set(str(self.config.get(chave) or ""))
        # Textos de observação: os mesmos da tela Parâmetros.
        for tipo, controle in self.observacoes_manual.items():
            personalizado = str(self.config.get(CHAVES_CONFIG_OBSERVACAO[tipo]) or "").strip()
            controle["modo"].set("proprio" if personalizado else "padrao")
            controle["texto"].set(personalizado or TEXTO_PADRAO_OBSERVACAO[tipo])

    def _conferir_campo_manual(self, campo: str):
        """Mostra, ao lado do campo, como o dado vai ser usado (verde) ou o
        que precisa corrigir (vermelho) - a cada tecla digitada."""
        rotulo, dica = self.retornos_manual[campo]
        texto = self.vars_manual[campo].get()
        if not texto.strip():
            _mostrar_retorno_campo(rotulo, "dica", dica)
            return
        try:
            _valor, exibicao = padronizar_campo(campo, texto)
            _mostrar_retorno_campo(rotulo, "ok", f"Certo: {exibicao}")
        except ValueError as erro:
            _mostrar_retorno_campo(rotulo, "erro", str(erro))

    def _proximo_campo_manual(self, indice: int):
        if indice + 1 < len(CAMPOS_TELA_MANUAL):
            self.entradas_manual[CAMPOS_TELA_MANUAL[indice + 1][0]].focus_set()
        else:
            self._adicionar_item_manual()

    def _limpar_formulario_manual(self):
        for var in self.vars_manual.values():
            var.set("")
        self.entradas_manual["cpf"].focus_set()

    def _formulario_manual_tem_texto(self) -> bool:
        return any(var.get().strip() for var in self.vars_manual.values())

    # ---- lista de pessoas ----
    @staticmethod
    def _observacao_item_manual(item: dict) -> str:
        """Para quem está voltando de uma tentativa anterior: de qual
        etapa o programa vai continuar (as já feitas não são refeitas)."""
        feitas = [etapa for etapa in ETAPAS if item.get(etapa.lower())]
        if not feitas:
            return ""
        proxima = next((etapa for etapa in ETAPAS if not item.get(etapa.lower())), "")
        return f"Já tem {', '.join(feitas)} - continua da {proxima}"

    def _todos_itens_manual(self):
        return list(self.itens_manual_pendentes) + list(self.itens_manual_resultados.items())

    def _adicionar_item_manual(self):
        textos = {campo: var.get() for campo, var in self.vars_manual.items()}
        item, erros = montar_item(textos)
        if erros:
            linhas = [
                f"- {TITULOS_CAMPOS_MANUAL[campo]}: {erros[campo]}"
                for campo, _titulo, _dica in CAMPOS_TELA_MANUAL if campo in erros
            ]
            messagebox.showwarning(
                "Amigo",
                "Ainda tem campo para corrigir (aparece com ❗ embaixo dele):\n\n" + "\n".join(linhas),
            )
            primeiro = next(campo for campo, _t, _d in CAMPOS_TELA_MANUAL if campo in erros)
            self.entradas_manual[primeiro].focus_set()
            return

        repetido = next(
            (
                numero for numero, outro in self._todos_itens_manual()
                if outro.get("situacao") != "simulado"  # simulação não gerou nada: não conta como repetido
                and (outro["cpf"], outro["ne"], outro["valor"]) == (item["cpf"], item["ne"], item["valor"])
            ),
            None,
        )
        if repetido is not None and not messagebox.askyesno(
            "Amigo",
            f"A pessoa nº {repetido} já tem esse mesmo CPF, Nota de Empenho e valor.\n\n"
            "Adicionar de novo mesmo assim? (isso vai gerar documentos em dobro no SIGEF)",
            icon="warning",
        ):
            return

        self.contador_manual += 1
        self.itens_manual_pendentes.append((self.contador_manual, item))
        log_info(
            f"Pessoa {self.contador_manual} adicionada à lista do Lançamento Manual "
            f"(CPF {item['exibicao']['cpf']}, valor {item['exibicao']['valor']})."
        )
        self._atualizar_lista_manual()
        self._limpar_formulario_manual()

    def _atualizar_lista_manual(self):
        tabela = self.tabela_pendentes_manual
        tabela.delete(*tabela.get_children())
        for numero, item in self.itens_manual_pendentes:
            exibicao = item["exibicao"]
            tabela.insert("", "end", iid=str(numero), values=(
                numero, exibicao["cpf"], exibicao["ne"], exibicao["banco"], exibicao["agencia"],
                exibicao["conta"], exibicao["valor"], self._observacao_item_manual(item),
            ))
        quantidade = len(self.itens_manual_pendentes)
        if quantidade:
            self.rotulo_lista_manual.config(
                text=f"👥  {quantidade} pessoa(s) esperando para serem lançadas. Para mudar alguém, "
                     "clique na linha dela e use os botões abaixo da tabela."
            )
        else:
            self.rotulo_lista_manual.config(
                text="📭  A lista está vazia. Preencha os dados acima e clique em \"➕ Adicionar à lista\"."
            )
        self._atualizar_botoes_manual()

    def _pendente_marcado_manual(self):
        selecao = self.tabela_pendentes_manual.selection()
        if not selecao:
            messagebox.showinfo("Amigo", "Primeiro clique na linha da pessoa, na tabela da lista.")
            return None
        numero = int(selecao[0])
        return next(((n, item) for n, item in self.itens_manual_pendentes if n == numero), None)

    def _corrigir_item_manual(self):
        marcado = self._pendente_marcado_manual()
        if marcado is None:
            return
        numero, item = marcado
        if self._observacao_item_manual(item):
            messagebox.showinfo(
                "Amigo",
                f"A pessoa nº {numero} já tem documento gerado no SIGEF "
                f"({self._observacao_item_manual(item)}).\n\n"
                "Para não gerar documentos errados ou em dobro, os dados dela não podem ser "
                "mudados aqui - só dá para tirar a pessoa da lista.",
            )
            return
        if self._formulario_manual_tem_texto() and not messagebox.askyesno(
            "Amigo", "O que está digitado no formulário agora vai ser trocado pelos dados dessa pessoa. Continuar?"
        ):
            return
        self.itens_manual_pendentes = [(n, i) for n, i in self.itens_manual_pendentes if n != numero]
        for campo, var in self.vars_manual.items():
            var.set(item["exibicao"][campo])
        self._atualizar_lista_manual()
        self.entradas_manual["cpf"].focus_set()
        log_info(f"Pessoa {numero} voltou para o formulário para ser corrigida.")

    def _remover_item_manual(self):
        marcado = self._pendente_marcado_manual()
        if marcado is None:
            return
        numero, item = marcado
        if not messagebox.askyesno(
            "Amigo", f"Tirar da lista a pessoa nº {numero} (CPF {item['exibicao']['cpf']})?"
        ):
            return
        self.itens_manual_pendentes = [(n, i) for n, i in self.itens_manual_pendentes if n != numero]
        self._atualizar_lista_manual()

    def _esvaziar_lista_manual(self):
        if not self.itens_manual_pendentes:
            return
        if not messagebox.askyesno("Amigo", "Tirar TODAS as pessoas da lista?"):
            return
        self.itens_manual_pendentes = []
        self._atualizar_lista_manual()

    def _atualizar_botoes_manual(self):
        """Liga/desliga os botões da aba conforme o momento: nada que mexa
        na lista de resultados enquanto os lançamentos estão rodando, e
        "Começar" desligado se QUALQUER automação (desta aba ou da
        planilha) estiver usando o navegador do SIGEF."""
        if not hasattr(self, "botao_comecar_manual"):
            return
        rodando = self.automacao_em_andamento
        tem_pendentes = bool(self.itens_manual_pendentes)
        tem_resultados = bool(self.itens_manual_resultados)
        rodando_aqui = self._manual_rodando

        def estado(ligado):
            return "normal" if ligado else "disabled"

        self.botao_comecar_manual.config(state=estado(not rodando and tem_pendentes))
        for botao in (self.botao_corrigir_manual, self.botao_remover_manual, self.botao_esvaziar_manual):
            botao.config(state=estado(tem_pendentes))
        self.botao_copiar_manual.config(state=estado(tem_resultados))
        self.botao_repetir_manual.config(state=estado(tem_resultados and not rodando_aqui))
        self.botao_limpar_resultados_manual.config(state=estado(tem_resultados and not rodando_aqui))

    # ---- execução ----
    def _comecar_lancamentos_manual(self):
        if self.automacao_em_andamento:
            messagebox.showinfo("Amigo", "Espere a automação que já está rodando terminar.")
            return
        if not self.itens_manual_pendentes:
            messagebox.showinfo(
                "Amigo", "A lista está vazia. Preencha os dados e clique em \"Adicionar à lista\" primeiro."
            )
            return
        parametros = self._parametros_manuais_da_tela()
        if parametros is None:
            return
        if self._formulario_manual_tem_texto() and not messagebox.askyesno(
            "Amigo",
            "Tem dados digitados no formulário que NÃO foram adicionados à lista - eles NÃO "
            "vão ser lançados.\n\nComeçar mesmo assim, só com quem já está na lista?",
        ):
            return

        quantidade = len(self.itens_manual_pendentes)
        simulacao = self.var_simulacao.get()
        if simulacao:
            abertura = (
                f"🧪 SIMULAR {quantidade} pessoa(s) no SIGEF.\n\n"
                "Para cada pessoa, a CE é preenchida e conferida (credor e valor), mas o programa "
                "PARA ANTES de confirmar - nada é gerado. Como a NL precisa de uma CE de verdade, "
                "a simulação de cada pessoa termina na CE.\n\n"
            )
        else:
            abertura = (
                f"Vai lançar {quantidade} pessoa(s) no SIGEF DE VERDADE (ambiente de produção).\n\n"
                "Para cada pessoa, nesta ordem: CE, depois NL, depois PP e por último OB - "
                "uma pessoa de cada vez.\n\n"
            )
        if not messagebox.askyesno(
            "Amigo",
            abertura +
            f"Data: {padronizar_parametro_manual('data', parametros['data'])[1]}\n"
            f"Processo: {parametros['processo']}\n"
            f"Mês referência: {parametros['mes_referencia']}\n"
            f"{aviso_mes_referencia(parametros['data'], parametros['mes_referencia'])}\n\n"
            "O navegador do SIGEF precisa estar aberto e com o login feito. Não mexa nele "
            "enquanto o programa trabalha.\n\nComeçar agora?",
        ):
            return

        fila = list(self.itens_manual_pendentes)
        self.itens_manual_pendentes = []
        for numero, item in fila:
            self.itens_manual_resultados[numero] = item
            item["situacao"] = "fila"
            self._desenhar_resultado_manual(numero, item)

        self.parar_manual.clear()
        execucao.iniciar_execucao(simulacao=simulacao)
        self.automacao_em_andamento = True
        self._manual_rodando = True
        self._atualizar_status_lateral()
        self._definir_botoes_automacao(False)  # a planilha espera o SIGEF ficar livre
        self.botao_parar_manual.config(state="normal", text="⏸  Parar depois da pessoa atual")
        self.barra_manual.config(maximum=len(fila) * len(ETAPAS), value=0)
        self._atualizar_lista_manual()
        log_info(
            f"Lançamento Manual: começando {'a SIMULAÇÃO de ' if simulacao else ''}{len(fila)} "
            f"pessoa(s), uma de cada vez."
        )

        # Os parâmetros usados ficam guardados para a próxima vez (na chave
        # própria desta aba) - e as automações recebem uma CÓPIA da
        # configuração com eles, sem alterar a da planilha.
        self._guardar_parametros_manual(parametros)
        config_lancamento = config_para_lancamento_manual(self.config, parametros)
        log_info(
            f"Parâmetros do Lançamento Manual: Data {parametros['data']} | Processo "
            f"{parametros['processo']} | Mês referência {parametros['mes_referencia']}."
        )
        threading.Thread(
            target=self._rodar_lancamentos_manual, args=(fila, config_lancamento), daemon=True
        ).start()

    def _rodar_lancamentos_manual(self, fila, config):
        """Roda em segundo plano (a janela continua respondendo). Toda
        mudança na tela volta para a thread da janela via `root.after`."""
        total = len(fila)
        nao_iniciados = []
        for posicao, (numero, item) in enumerate(fila, start=1):
            if self.parar_manual.is_set():
                nao_iniciados = fila[posicao - 1:]
                break

            def ao_iniciar(etapa, posicao=posicao, numero=numero, item=item):
                self.root.after(0, lambda: self._etapa_manual_iniciou(posicao, total, numero, item, etapa))

            def ao_terminar(etapa, ok, texto, numero=numero, item=item):
                self.root.after(0, lambda: self._etapa_manual_terminou(numero, item, etapa, ok))

            try:
                executar_cadeia(item, numero, config, ao_iniciar, ao_terminar)
            except Exception as erro:  # rede de segurança: nunca deixa a fila travada
                item.update({"situacao": "parou", "motivo": f"Erro inesperado: {erro}"})
                item["etapa_parada"] = item.get("etapa_parada") or next(
                    (etapa for etapa in ETAPAS if not item.get(etapa.lower())), ""
                )
                log_erro(f"Pessoa {numero}: erro inesperado no Lançamento Manual: {erro}")

            self.root.after(0, lambda p=posicao, n=numero, i=item: self._pessoa_manual_terminou(p, n, i))

        self.root.after(0, lambda: self._finalizar_lancamentos_manual(fila, nao_iniciados))

    def _etapa_manual_iniciou(self, posicao, total, numero, item, etapa):
        item["etapa_atual"] = etapa
        _configurar(self.rotulo_andamento_manual, 
            text=f"⏳  Pessoa {posicao} de {total} (nº {numero}, CPF {item['exibicao']['cpf']}): "
                 f"fazendo a {NOMES_ETAPAS[etapa]} - etapa {ETAPAS.index(etapa) + 1} de 4.\n"
                 "Aguarde e não mexa no navegador.",
            style="Subtitulo.TLabel",
        )
        self._desenhar_resultado_manual(numero, item)

    def _etapa_manual_terminou(self, numero, item, etapa, ok):
        item["etapa_atual"] = ""
        if ok:
            self.barra_manual.step(1)
        self._desenhar_resultado_manual(numero, item)

    def _pessoa_manual_terminou(self, posicao, numero, item):
        item["etapa_atual"] = ""
        self.barra_manual.config(value=posicao * len(ETAPAS))
        self._desenhar_resultado_manual(numero, item)

    def _pedir_parada_manual(self):
        self.parar_manual.set()
        self.botao_parar_manual.config(state="disabled", text="⏳  Vai parar depois da pessoa atual...")
        log_info("Lançamento Manual: vai parar assim que a pessoa atual terminar.")

    def _finalizar_lancamentos_manual(self, fila, nao_iniciados):
        # Quem não chegou a começar (botão "Parar") volta para a lista.
        for numero, item in nao_iniciados:
            self.itens_manual_resultados.pop(numero, None)
            if self.tabela_resultados_manual.exists(str(numero)):
                self.tabela_resultados_manual.delete(str(numero))
            item["situacao"] = "pendente"
            self.itens_manual_pendentes.append((numero, item))

        simulacao = execucao.em_simulacao()
        execucao.finalizar_execucao()
        self.automacao_em_andamento = False
        self._manual_rodando = False
        self._atualizar_status_lateral()
        self.botao_parar_manual.config(state="disabled", text="⏸  Parar depois da pessoa atual")
        if self.dados is not None and self.worksheet is not None:
            self._definir_botoes_automacao(True)

        numeros_nao_iniciados = {numero for numero, _item in nao_iniciados}
        feitos = [item for numero, item in fila if numero not in numeros_nao_iniciados]
        concluidos = sum(1 for item in feitos if item["situacao"] == "concluido")
        simulados = sum(1 for item in feitos if item["situacao"] == "simulado")
        parados = len(feitos) - concluidos - simulados

        if simulacao:
            resumo = f"Simulação terminada! {simulados} pessoa(s) conferida(s) sem problema - nada foi gerado"
        else:
            resumo = f"Terminou! {concluidos} pessoa(s) concluída(s)"
        if parados:
            resumo += f", {parados} parou(aram) em alguma etapa"
        if nao_iniciados:
            resumo += f", {len(nao_iniciados)} voltou(aram) para a lista sem começar"
        resumo += "."
        _configurar(self.rotulo_andamento_manual, 
            text=("🎉  " if not parados else "⚠  ") + resumo + " Os números estão em \"⑤ Resultados\", logo abaixo.",
            style="Sucesso.TLabel" if not parados else "Aviso.TLabel",
        )
        self._atualizar_lista_manual()
        log_sucesso(f"Lançamento Manual: {resumo}")

        mensagem = resumo + "\n\n"
        if parados:
            mensagem += (
                "Para ver por que alguém parou, clique na linha vermelha da tabela "
                "\"Resultados\". O botão \"Tentar de novo as que pararam\" continua da etapa "
                "em que a pessoa parou, sem refazer as que já deram certo.\n\n"
            )
        mensagem += (
            "LEMBRE: os números NÃO ficam salvos. Anote ou clique em \"Copiar resultados\" "
            "e cole (Ctrl+V) no Excel, no Word ou no bloco de notas."
        )
        messagebox.showinfo("Amigo", mensagem)

    # ---- resultados ----
    def _desenhar_resultado_manual(self, numero, item):
        """Cria/atualiza a linha da pessoa na tabela "Resultados"."""
        etapa_atual = item.get("etapa_atual", "")

        def celula(etapa):
            valor = item.get(etapa.lower())
            if valor and execucao.eh_simulado(valor):
                return "🧪 simulado ✓"
            if valor:
                return valor
            if etapa == etapa_atual:
                return "⏳ fazendo..."
            return "-"

        situacao = item.get("situacao")
        if situacao == "concluido":
            texto_situacao, marca = "✅ Concluído", "concluido"
        elif situacao == "simulado":
            texto_situacao, marca = "🧪 Simulado", "andamento"
        elif situacao == "parou":
            texto_situacao, marca = f"❗ Parou na {item.get('etapa_parada')}", "parou"
        elif situacao == "andamento":
            texto_situacao, marca = "⏳ Em andamento", "andamento"
        else:
            texto_situacao, marca = "🕒 Na fila", ""

        valores = (
            numero, item["exibicao"]["cpf"], item["exibicao"]["valor"],
            celula("CE"), celula("NL"), celula("PP"), celula("OB"), texto_situacao,
        )
        tabela = self.tabela_resultados_manual
        iid = str(numero)
        if tabela.exists(iid):
            tabela.item(iid, values=valores, tags=(marca,))
        else:
            tabela.insert("", "end", iid=iid, values=valores, tags=(marca,))
        tabela.see(iid)
        if tabela.selection() == (iid,):
            self._mostrar_detalhe_resultado_manual()
        self._atualizar_botoes_manual()

    def _mostrar_detalhe_resultado_manual(self):
        selecao = self.tabela_resultados_manual.selection()
        if not selecao:
            _configurar(self.rotulo_detalhe_manual, text="", style="Instrucao.TLabel")
            return
        numero = int(selecao[0])
        item = self.itens_manual_resultados.get(numero)
        if item is None:
            return
        exibicao = item["exibicao"]
        texto = (
            f"👤  Pessoa nº {numero}  •  CPF {exibicao['cpf']}  •  NE {exibicao['ne']}  •  {exibicao['valor']}\n"
            f"🏦  Banco {exibicao['banco']}  •  Agência {exibicao['agencia']}  •  Conta {exibicao['conta']}\n"
            f"📄  CE {item['ce'] or '-'}   →   NL {item['nl'] or '-'}   →   PP {item['pp'] or '-'}   →   OB {item['ob'] or '-'}"
        )
        estilo = "Instrucao.TLabel"
        if item["situacao"] == "parou":
            texto += (
                f"\n❗  Parou na etapa {NOMES_ETAPAS.get(item['etapa_parada'], item['etapa_parada'])}. "
                f"Motivo: {item['motivo']}\n"
                "💡  Resolva o problema (no SIGEF ou nos dados) e clique em \"🔁 Tentar de novo as que "
                "pararam\" - o programa continua desta etapa, sem refazer as anteriores."
            )
            estilo = "Erro.TLabel"
        elif item["situacao"] == "simulado":
            texto += (
                f"\n🧪  {item['motivo']}\n"
                "💡  Para lançar de verdade, desligue o modo simulação (menu da esquerda) e "
                "adicione a pessoa de novo à lista."
            )
            estilo = "Aviso.TLabel"
        _configurar(self.rotulo_detalhe_manual, text=texto, style=estilo)

    def _copiar_resultados_manual(self):
        if not self.itens_manual_resultados:
            return
        texto = texto_resultados(sorted(self.itens_manual_resultados.items()))
        self.root.clipboard_clear()
        self.root.clipboard_append(texto)
        messagebox.showinfo(
            "Amigo",
            "Resultados copiados!\n\nAgora é só abrir o Excel, o Word ou o bloco de notas e "
            "colar (aperte Ctrl e V juntos).",
        )

    def _repetir_itens_parados_manual(self):
        parados = [(n, i) for n, i in sorted(self.itens_manual_resultados.items()) if i["situacao"] == "parou"]
        if not parados:
            messagebox.showinfo("Amigo", "Nenhuma pessoa parou - não há o que tentar de novo.")
            return
        if not messagebox.askyesno(
            "Amigo",
            f"{len(parados)} pessoa(s) vão voltar para a lista. Cada uma continua da etapa em que "
            "parou (o que já foi gerado não é feito de novo).\n\nDepois é só clicar em "
            "\"Começar lançamentos\". Continuar?",
        ):
            return
        for numero, item in parados:
            del self.itens_manual_resultados[numero]
            self.tabela_resultados_manual.delete(str(numero))
            item.update({"situacao": "pendente", "etapa_parada": "", "motivo": ""})
            self.itens_manual_pendentes.append((numero, item))
        self.itens_manual_pendentes.sort(key=lambda par: par[0])
        _configurar(self.rotulo_detalhe_manual, text="", style="Instrucao.TLabel")
        self._atualizar_lista_manual()

    def _limpar_resultados_manual(self):
        if not self.itens_manual_resultados:
            return
        if not messagebox.askyesno(
            "Amigo",
            "Os números vão sumir da tela e NÃO dá para recuperar (eles não ficam salvos).\n\n"
            "Você já anotou ou copiou os resultados?",
            icon="warning",
        ):
            return
        self.itens_manual_resultados = {}
        self.tabela_resultados_manual.delete(*self.tabela_resultados_manual.get_children())
        _configurar(self.rotulo_detalhe_manual, text="", style="Instrucao.TLabel")
        self._atualizar_botoes_manual()

    # ------------------------------------------------------------------
    # ABA: PARÂMETROS
    # ------------------------------------------------------------------
    def _montar_aba_parametros(self):
        conteudo = _criar_area_rolavel(self.aba_parametros)
        self.retornos_parametros = {}

        _criar_cabecalho_aba(
            conteudo,
            "⚙  Parâmetros da planilha",
            "Informações usadas quando você lança 🤖 com a planilha. Preencha e clique em "
            "\"💾 Salvar parâmetros\". (O ✍ Sem planilha tem parâmetros próprios.)",
            passos=[
                "Preencha a data, o processo e o mês referência do pagamento.",
                "Clique em \"📂 Procurar\" e escolha o arquivo do Excel com as pessoas.",
                "Diga em qual linha do Excel está a primeira pessoa (\"Linha inicial\").",
                "Clique em \"💾 Salvar parâmetros\". Pronto: volte para 🤖 Com planilha.",
            ],
        )

        def campo(pai, linha, coluna, chave, titulo, dica, largura=22):
            var = tk.StringVar(value=str(self.config.get(chave, "") or ""))
            _entrada, retorno = self._montar_campo(pai, linha, coluna, titulo, dica, var, largura=largura)
            var.trace_add("write", lambda *_args: self._conferir_parametros())
            self.vars_parametros[chave] = var
            self.retornos_parametros[chave] = (retorno, dica)

        # ---- Dados do pagamento ----
        frame_pagamento = Cartao(conteudo, text="📝  Dados do pagamento", padding=14)
        frame_pagamento.pack(fill="x", pady=(0, 22))
        campo(frame_pagamento, 0, 0, "data", "📅  Data", "Ex: 30062026 ou 30/06/2026", 18)
        campo(frame_pagamento, 0, 1, "processo", "📁  Processo", "Processo do Amigo Voluntário. Ex: 0029.037004/2026-72", 26)
        campo(frame_pagamento, 0, 2, "mes_referencia", "🗓  Mês referência", "Ex: 06/2026 - sempre o mês anterior", 18)
        campo(frame_pagamento, 1, 0, "ano_sigef", "🏛  Ano do SIGEF (exercício)",
              "Vazio = automático, pelo ano da Data", 10)

        # ---- Textos de observação (padrão ou próprio) ----
        frame_observacoes = Cartao(conteudo, text="📝  Textos de observação")
        frame_observacoes.pack(fill="x", pady=(0, 22))
        self.observacoes_planilha = self._montar_observacoes(
            frame_observacoes, self.config,
            lambda: (self.vars_parametros["mes_referencia"].get().strip(),
                     self.vars_parametros["processo"].get().strip()),
        )

        # ---- Planilha ----
        frame_excel = Cartao(conteudo, text="📊  Planilha do Excel", padding=14)
        frame_excel.pack(fill="x", pady=(0, 22))

        bloco_arquivo = ttk.Frame(frame_excel)
        bloco_arquivo.grid(row=0, column=0, columnspan=3, sticky="we", pady=6)
        ttk.Label(bloco_arquivo, text="📂  Arquivo do Excel (.xlsx)", style="Campo.TLabel").pack(anchor="w")
        linha_arquivo = ttk.Frame(bloco_arquivo)
        linha_arquivo.pack(fill="x", pady=(4, 3))
        var_planilha = tk.StringVar(value=str(self.config.get("caminho_planilha", "") or ""))
        var_planilha.trace_add("write", lambda *_args: self._conferir_parametros())
        self.vars_parametros["caminho_planilha"] = var_planilha
        ttk.Entry(linha_arquivo, textvariable=var_planilha).pack(side="left", fill="x", expand=True)
        botao_procurar = ttk.Button(linha_arquivo, text="📂  Procurar", command=self._procurar_planilha)
        botao_procurar.pack(side="left", padx=(8, 0))
        Dica(botao_procurar, "Abre uma janela para você escolher o arquivo do Excel no computador.")
        dica_arquivo = "Clique em 📂 Procurar e escolha a planilha"
        retorno_arquivo = _texto_ajuda_campo(bloco_arquivo, "")
        retorno_arquivo.config(wraplength=900)
        retorno_arquivo.pack(anchor="w")
        self.retornos_parametros["caminho_planilha"] = (retorno_arquivo, dica_arquivo)
        frame_excel.columnconfigure(2, weight=1)

        campo(frame_excel, 1, 0, "aba_planilha", "📑  Nome da aba", "Vazio = usa a primeira aba", 22)
        campo(frame_excel, 1, 1, "linha_inicial", "🔢  Linha inicial", "Linha do Excel onde está a 1ª pessoa. Ex: 37", 10)

        botao_salvar = ttk.Button(
            conteudo, text="💾  Salvar parâmetros", style="Principal.TButton",
            command=self._salvar_parametros,
        )
        botao_salvar.pack(anchor="w", padx=2, pady=(4, 12))
        self._conferir_parametros()

    def _conferir_parametros(self):
        """Mesmo retorno visual do Lançamento Manual (💡 dica, ✅ certo, ⚠
        atenção, ❗ corrigir), a cada tecla digitada - só mostra, não muda
        nada; a gravação continua sendo no botão \"Salvar parâmetros\"."""
        if not getattr(self, "retornos_parametros", None):
            return
        textos = {chave: var.get() for chave, var in self.vars_parametros.items()}
        for tipo, controle in getattr(self, "observacoes_planilha", {}).items():
            self._atualizar_observacao(
                controle, tipo,
                lambda: (textos.get("mes_referencia", "").strip(), textos.get("processo", "").strip()),
            )
        for chave, (rotulo, dica) in self.retornos_parametros.items():
            texto = textos.get(chave, "").strip()
            if not texto:
                if chave == "ano_sigef":
                    ano = ano_do_exercicio({"data": textos.get("data", "")})
                    dica = f"Vazio = automático: vai usar o SIGEF{ano} (ano da Data)"
                _mostrar_retorno_campo(rotulo, "dica", dica)
                continue
            if chave in ("data", "processo", "mes_referencia"):
                try:
                    _valor, exibicao = padronizar_parametro_manual(chave, texto)
                except ValueError as erro:
                    _mostrar_retorno_campo(rotulo, "erro", str(erro))
                    continue
                aviso = aviso_mes_referencia(textos.get("data", ""), texto) if chave == "mes_referencia" else ""
                if aviso:
                    _mostrar_retorno_campo(rotulo, "aviso", f"{exibicao}. {aviso}")
                else:
                    _mostrar_retorno_campo(rotulo, "ok", f"Certo: {exibicao}")
            elif chave == "ano_sigef":
                ano_da_data = ano_do_exercicio({"data": textos.get("data", "")})
                if texto.isdigit() and len(texto) == 4 and 2000 <= int(texto) <= 2100:
                    if int(texto) != ano_da_data:
                        _mostrar_retorno_campo(
                            rotulo, "aviso",
                            f"Vai usar o SIGEF{texto}, diferente do ano da Data ({ano_da_data}). "
                            f"Confira se é isso mesmo.",
                        )
                    else:
                        _mostrar_retorno_campo(rotulo, "ok", f"Vai usar o SIGEF{texto}")
                else:
                    _mostrar_retorno_campo(rotulo, "erro", "Use o ano com 4 números (ex: 2027), ou deixe vazio")
            elif chave == "linha_inicial":
                if texto.isdigit() and int(texto) >= 1:
                    _mostrar_retorno_campo(rotulo, "ok", f"A leitura começa na linha {int(texto)} do Excel")
                else:
                    _mostrar_retorno_campo(rotulo, "erro", "Use só o número da linha (ex: 37)")
            elif chave == "aba_planilha":
                _mostrar_retorno_campo(rotulo, "ok", f"Vai usar a aba \"{texto}\"")
            elif chave == "caminho_planilha":
                caminho = texto.strip('"')
                if os.path.isfile(caminho):
                    _mostrar_retorno_campo(rotulo, "ok", f"Arquivo encontrado: {os.path.basename(caminho)}")
                else:
                    _mostrar_retorno_campo(rotulo, "erro", "Não achei esse arquivo. Use o botão 📂 Procurar.")

    def _procurar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.vars_parametros["caminho_planilha"].set(caminho)

    def _salvar_parametros(self):
        linha_inicial_texto = self.vars_parametros["linha_inicial"].get().strip()
        if linha_inicial_texto and not linha_inicial_texto.isdigit():
            messagebox.showerror("Amigo", "A \"Linha inicial\" precisa ser só um número (ex: 37).")
            return

        # Data, processo e mês: se preenchidos, precisam estar num formato
        # que o SIGEF aceite - e são gravados JÁ padronizados (ex:
        # "30/06/2026" vira "30062026", "6/2026" vira "06/2026"), o mesmo
        # formato de sempre das automações. Vazio continua permitido.
        padronizados = {}
        erros = []
        titulos = {"data": "Data", "processo": "Processo", "mes_referencia": "Mês referência"}
        for chave, titulo in titulos.items():
            texto = self.vars_parametros[chave].get().strip()
            if not texto:
                padronizados[chave] = ""
                continue
            try:
                padronizados[chave], _ = padronizar_parametro_manual(chave, texto)
            except ValueError as erro:
                erros.append(f"- {titulo}: {erro}")
        ano_sigef = self.vars_parametros["ano_sigef"].get().strip()
        if ano_sigef and not (ano_sigef.isdigit() and len(ano_sigef) == 4 and 2000 <= int(ano_sigef) <= 2100):
            erros.append("- Ano do SIGEF: use o ano com 4 números (ex: 2027), ou deixe vazio.")
        valores_observacoes, erros_observacoes = self._valores_observacoes(self.observacoes_planilha)
        erros.extend(erros_observacoes)
        if erros:
            messagebox.showwarning("Amigo", "Corrija antes de salvar (o que está com ❗):\n\n" + "\n".join(erros))
            return
        self.config.update(valores_observacoes)

        for chave, var in self.vars_parametros.items():
            valor = var.get().strip()
            if chave == "linha_inicial":
                if valor:
                    self.config[chave] = int(valor)
            elif chave == "caminho_planilha":
                self.config[chave] = valor.strip('"')
            elif chave in padronizados:
                self.config[chave] = padronizados[chave]
            else:
                self.config[chave] = valor

        salvar_configuracoes(self.config)
        for chave, valor in padronizados.items():
            self.vars_parametros[chave].set(valor)
        log_sucesso("Parâmetros atualizados com sucesso!")

    # ------------------------------------------------------------------
    # ABA: COLUNAS RETRÁTEIS
    # ------------------------------------------------------------------
    def _montar_aba_colunas(self):
        # Barra de botões FIXA embaixo (sempre visível, mesmo rolando a lista).
        botoes = ttk.Frame(self.aba_colunas)
        botoes.pack(side="bottom", fill="x", pady=(10, 0))
        frame_interno = _criar_area_rolavel(self.aba_colunas)

        _criar_cabecalho_aba(
            frame_interno,
            "🧱  Colunas Retráteis",
            "Diz ao programa em qual coluna (letra) do Excel está cada informação. "
            "Se você não sabe o que é isso, não precisa mexer: já vem pronto.",
            passos=[
                "Abra sua planilha e veja a letra de cada coluna (A, B, C...).",
                "Escreva aqui a letra certa de cada informação. Passe o mouse na caixa para ver "
                "a explicação completa.",
                "Duas informações não podem ficar na mesma letra: se isso acontecer, a caixa fica "
                "vermelha e o botão de salvar trava até corrigir.",
                "Clique em \"💾 Salvar colunas\". Se errar, \"↩ Voltar ao padrão de fábrica\" "
                "desfaz tudo.",
            ],
        )

        colunas_atuais = self.config.get("colunas")
        if not isinstance(colunas_atuais, dict):
            colunas_atuais = dict(COLUNAS_PADRAO)

        grupos = [
            ("👤  Dados da pessoa e do pagamento", CAMPOS_COLUNAS_BASICAS),
            ("📄  Números gerados pelo SIGEF (CE → NL → PP → OB)", CAMPOS_COLUNAS_ENCADEADAS),
            ("📝  Mensagens e conferências", CAMPOS_COLUNAS_AVANCADAS),
        ]

        for titulo, campos in grupos:
            grupo_frame = Cartao(frame_interno, text=titulo, padding=12)
            grupo_frame.pack(fill="x", pady=(0, 18))
            for indice, (chave, rotulo, padrao) in enumerate(campos):
                # 2 colunas lado a lado: fica mais curto e fácil de comparar.
                linha, coluna = divmod(indice, 2)
                bloco = ttk.Frame(grupo_frame)
                bloco.grid(row=linha, column=coluna, sticky="w", padx=(0 if coluna == 0 else 40, 0), pady=5)
                ttk.Label(bloco, text=NOMES_CURTOS_COLUNAS.get(chave, rotulo), width=26).pack(side="left")
                valor_atual = str(colunas_atuais.get(chave) or padrao).strip().upper()
                var = tk.StringVar(value=valor_atual)
                var.trace_add("write", lambda *_args: self._revalidar_colunas())
                entrada = ttk.Entry(bloco, textvariable=var, width=5, justify="center")
                entrada.pack(side="left", padx=(8, 10))
                ttk.Label(bloco, text=f"padrão: {padrao}", style="Dica.TLabel").pack(side="left")
                Dica(entrada, f"{rotulo}\n\nPadrão de fábrica: coluna {padrao}.")
                self.vars_colunas[chave] = var
                self.entradas_colunas[chave] = entrada

        self.rotulo_colisao_colunas = ttk.Label(botoes, text="", style="Erro.TLabel", wraplength=560, justify="left")
        self.rotulo_colisao_colunas.pack(side="left")
        botao_restaurar = ttk.Button(botoes, text="↩  Voltar ao padrão de fábrica", command=self._restaurar_colunas_padrao)
        botao_restaurar.pack(side="right", padx=(0, 10))
        Dica(botao_restaurar, "Coloca de volta as letras originais (B, D, E, F, G, H, I, J, K, L, M, N). "
                              "Só vale depois de clicar em \"💾 Salvar colunas\".")
        self.botao_salvar_colunas = ttk.Button(
            botoes, text="💾  Salvar colunas", style="Principal.TButton", command=self._salvar_colunas,
        )
        self.botao_salvar_colunas.pack(side="right", padx=(0, 10))

        self._revalidar_colunas()

    def _colunas_propostas(self) -> dict:
        return {chave: var.get().strip().upper() for chave, var in self.vars_colunas.items()}

    def _revalidar_colunas(self):
        propostas = self._colunas_propostas()
        colisoes = _detectar_colisoes(propostas)

        chaves_em_conflito = set()
        for _letra, itens in colisoes.items():
            chaves_em_conflito.update(chave for chave, _rotulo in itens)

        for chave, entrada in self.entradas_colunas.items():
            estilo = ESTILO_ENTRADA_CONFLITO if chave in chaves_em_conflito else ESTILO_ENTRADA_NORMAL
            entrada.configure(style=estilo)

        if colisoes:
            partes = [
                f"coluna {letra}: {', '.join(rotulo for _c, rotulo in itens)}"
                for letra, itens in colisoes.items()
            ]
            self.rotulo_colisao_colunas.config(
                text="❗  Letra repetida - corrija antes de salvar: " + " | ".join(partes)
            )
            self.botao_salvar_colunas.config(state="disabled")
        else:
            self.rotulo_colisao_colunas.config(text="")
            self.botao_salvar_colunas.config(state="normal")

    def _salvar_colunas(self):
        propostas = self._colunas_propostas()
        if _detectar_colisoes(propostas):
            messagebox.showerror("Amigo", "Ainda há colunas duplicadas. Corrija antes de salvar.")
            return

        for chave, letra in propostas.items():
            if not letra:
                messagebox.showerror("Amigo", "Nenhuma coluna pode ficar vazia.")
                return

        self.config["colunas"] = propostas
        salvar_configuracoes(self.config)
        aplicar_colunas(propostas)
        log_sucesso("Colunas atualizadas com sucesso!")

    def _restaurar_colunas_padrao(self):
        resposta = messagebox.askyesno(
            "Amigo", "Isso volta TODAS as colunas para o padrão de fábrica. Confirma?"
        )
        if not resposta:
            return
        for chave, var in self.vars_colunas.items():
            var.set(COLUNAS_PADRAO[chave])
        # As mudanças acima só ficam valendo de verdade quando o usuário
        # clicar em "Salvar colunas" - mesma confirmação em duas etapas do
        # menu de terminal.

    # ------------------------------------------------------------------
    # ABA: ESTATÍSTICAS
    #
    # Ao lado da aba "Colunas Retráteis": tempo economizado com a
    # automação (frente a uma estimativa de tempo manual, ajustável nesta
    # própria aba) e como esse tempo se comporta em função do número de
    # itens a processar - dados pensados para embasar um estudo de
    # implementação de automações em larga escala na secretaria, e não só
    # o uso individual do Amigo. Os números vêm do mesmo historico.json da
    # aba "Relatório" (via `amigo.estatisticas`), então nenhuma automação
    # precisa gravar nada diferente do que já grava hoje.
    # ------------------------------------------------------------------
    def _montar_aba_estatisticas(self):
        conteudo = _criar_area_rolavel(self.aba_estatisticas)

        _criar_cabecalho_aba(
            conteudo,
            "📊  Estatísticas",
            "Quanto tempo a automação já economizou, comparado a fazer tudo à mão - e quanto "
            "tempo levaria se o volume crescesse.",
            passos=[
                "Escolha o período: \"Todo o período\" ou um mês.",
                "O tempo da automação é medido de verdade, sozinho, a cada execução.",
                "O tempo \"à mão\" é uma estimativa: ajuste os minutos por item conforme a "
                "realidade do seu setor e clique em \"💾 Salvar estimativas\".",
                "Veja nos quadros o total economizado e, embaixo, a projeção para volumes maiores "
                "(útil para planejar o uso em toda a secretaria).",
            ],
        )

        # ---- período ----
        topo = ttk.Frame(conteudo)
        topo.pack(fill="x", pady=(0, 22))
        ttk.Label(topo, text="📆  Período:", style="Subtitulo.TLabel").pack(side="left", padx=(0, 10))
        self.var_periodo_estatisticas = tk.StringVar()
        self.combo_periodo_estatisticas = ttk.Combobox(
            topo, textvariable=self.var_periodo_estatisticas, state="readonly", width=16,
        )
        self.combo_periodo_estatisticas.pack(side="left")
        self.combo_periodo_estatisticas.bind(
            "<<ComboboxSelected>>", lambda _evento: self._atualizar_estatisticas()
        )
        ttk.Button(topo, text="🔄  Atualizar", command=self._atualizar_estatisticas).pack(side="left", padx=(10, 0))

        # ---- tempo manual estimado (editável) ----
        frame_tempos = Cartao(
            conteudo, text="✋  Quanto tempo leva fazer 1 item à mão (sem o Amigo)", padding=12,
        )
        frame_tempos.pack(fill="x", pady=(0, 22))
        tempos_atuais = obter_tempos_manuais(self.config)
        for coluna, categoria in enumerate(CATEGORIAS_ESTATISTICAS):
            bloco_categoria = ttk.Frame(frame_tempos)
            bloco_categoria.grid(row=0, column=coluna, padx=(0 if coluna == 0 else 16, 0), sticky="w")
            ttk.Label(bloco_categoria, text=categoria, style="Campo.TLabel").pack(anchor="w")
            var = tk.StringVar(value=f"{tempos_atuais.get(categoria, 0):.1f}")
            entrada_tempo = ttk.Entry(bloco_categoria, textvariable=var, width=8, justify="center")
            entrada_tempo.pack(anchor="w", pady=(2, 0))
            Dica(entrada_tempo, f"Quantos minutos uma pessoa leva para fazer 1 {categoria} à mão no SIGEF "
                                "(abrir a tela, digitar, conferir e confirmar). Pode usar vírgula: 2,5.")
            ttk.Label(bloco_categoria, text="minutos", style="Dica.TLabel").pack(anchor="w")
            self.vars_tempos_manuais[categoria] = var
        frame_tempos.columnconfigure(len(CATEGORIAS_ESTATISTICAS), weight=1)
        ttk.Button(
            frame_tempos, text="💾  Salvar estimativas", command=self._salvar_tempos_manuais,
        ).grid(row=0, column=len(CATEGORIAS_ESTATISTICAS), padx=(20, 0), sticky="e")

        # ---- cartões-resumo (totais do período) ----
        frame_cartoes = ttk.Frame(conteudo)
        frame_cartoes.pack(fill="x", pady=(0, 22))
        # Chave interna (usada em `_atualizar_estatisticas`) -> nome curto na tela.
        titulos_cartoes = {
            "Itens processados": "📦  Itens feitos", "Tempo com automação": "🤖  Com automação",
            "Tempo manual estimado": "✋  À mão (estimado)", "Tempo economizado": "🎉  Economizado",
        }
        for indice, titulo in enumerate(
            ["Itens processados", "Tempo com automação", "Tempo manual estimado", "Tempo economizado"]
        ):
            cartao = Cartao(frame_cartoes)
            cartao.grid(row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 16, 0))
            frame_cartoes.columnconfigure(indice, weight=1, uniform="kpi")
            ttk.Label(cartao, text=titulos_cartoes[titulo], style="Campo.TLabel").pack(anchor="w", pady=(0, 8))
            rotulo_valor = ttk.Label(cartao, text="-", style="NumeroGrande.TLabel")
            rotulo_valor.pack(anchor="w")
            rotulo_extra = ttk.Label(cartao, text="", style="Dica.TLabel", wraplength=230, justify="left")
            rotulo_extra.pack(anchor="w", pady=(6, 0))
            self.cartoes_estatisticas[titulo] = (rotulo_valor, rotulo_extra)

        # ---- detalhamento por categoria ----
        frame_categorias = Cartao(conteudo, text="🧾  Economia por tipo de automação", padding=12)
        frame_categorias.pack(fill="x", pady=(0, 22))
        cabecalho_categorias = ttk.Frame(frame_categorias)
        cabecalho_categorias.pack(fill="x")
        for coluna, titulo in enumerate(
            ["Categoria", "Itens", "Com automação", "À mão (estim.)", "Economizado", "Economia"]
        ):
            ttk.Label(
                cabecalho_categorias, text=titulo, style="Campo.TLabel", width=15,
            ).grid(row=0, column=coluna, sticky="w", padx=4, pady=(0, 8))
        self.frame_linhas_estatisticas = ttk.Frame(frame_categorias)
        self.frame_linhas_estatisticas.pack(fill="x")

        # ---- projeção para larga escala ----
        frame_projecao = Cartao(
            conteudo,
            text="🚀  E se o volume crescesse? (projeção)",
            padding=12,
        )
        frame_projecao.pack(fill="x", pady=(0, 22))
        ttk.Label(
            frame_projecao,
            text="Estimativa para volumes maiores (por exemplo, a secretaria inteira usando), "
                 "mantendo a velocidade média já medida nas execuções reais do período escolhido.",
            style="Dica.TLabel", wraplength=1000, justify="left",
        ).pack(anchor="w", pady=(0, 10))
        cabecalho_projecao = ttk.Frame(frame_projecao)
        cabecalho_projecao.pack(fill="x")
        for coluna, titulo in enumerate(
            ["Itens", "Tempo automatizado", "Tempo manual (est.)", "Tempo economizado"]
        ):
            ttk.Label(
                cabecalho_projecao, text=titulo, style="Campo.TLabel", width=20,
            ).grid(row=0, column=coluna, sticky="w", padx=4, pady=(0, 8))
        self.frame_linhas_projecao = ttk.Frame(frame_projecao)
        self.frame_linhas_projecao.pack(fill="x")

        self._atualizar_estatisticas()

    def _salvar_tempos_manuais(self):
        novos_tempos = {}
        for categoria, var in self.vars_tempos_manuais.items():
            texto = var.get().strip().replace(",", ".")
            try:
                novos_tempos[categoria] = max(0.0, float(texto)) if texto else 0.0
            except ValueError:
                messagebox.showerror(
                    "Amigo",
                    f"O tempo manual de '{categoria}' precisa ser um número (ex: 3 ou 3.5).",
                )
                return
        salvar_tempos_manuais(self.config, novos_tempos)
        log_sucesso("Estimativas de tempo manual atualizadas.")
        self._atualizar_estatisticas()

    def _popular_periodos_estatisticas(self):
        """Preenche o seletor de período com 'Todo o período' + os meses
        que já têm alguma execução registrada - mesmo esquema do seletor
        de mês da aba Relatório (`_popular_meses_relatorio`), preservando
        a seleção do usuário quando ela continua existindo na lista."""
        selecao_atual = self.var_periodo_estatisticas.get()
        periodos = ["Todo o período"] + meses_disponiveis()
        self.combo_periodo_estatisticas["values"] = periodos
        if selecao_atual in periodos:
            self.var_periodo_estatisticas.set(selecao_atual)
        else:
            self.var_periodo_estatisticas.set(periodos[0])

    def _atualizar_estatisticas(self):
        """Recarrega a aba Estatísticas a partir do histórico em disco,
        para o período escolhido. Chamada ao abrir a aba, ao trocar de
        período, no botão 'Atualizar' e automaticamente assim que
        qualquer automação termina (ver `_executar_automacao`) - mesmo
        padrão da aba Relatório (`_atualizar_relatorio`)."""
        if not hasattr(self, "var_periodo_estatisticas"):
            return  # aba ainda não foi montada

        self._popular_periodos_estatisticas()
        periodo = self.var_periodo_estatisticas.get() or "Todo o período"
        dados = calcular_estatisticas(periodo, self.config)
        totais = dados["totais"]

        def texto_percentual(valor):
            sinal = "+" if valor >= 0 else ""
            return f"{sinal}{valor:.0f}%"

        if "Itens processados" in self.cartoes_estatisticas:
            rotulo_valor, rotulo_extra = self.cartoes_estatisticas["Itens processados"]
            rotulo_valor.config(text=str(totais["processados"]))
            rotulo_extra.config(text=f"{totais['execucoes']} execução(ões)")
        if "Tempo com automação" in self.cartoes_estatisticas:
            rotulo_valor, rotulo_extra = self.cartoes_estatisticas["Tempo com automação"]
            rotulo_valor.config(text=formatar_duracao(totais["duracao_automacao_segundos"]))
            rotulo_extra.config(
                text=(
                    f"{formatar_duracao(totais['segundos_por_item_automacao'])}/item"
                    if totais["processados"] else "-"
                )
            )
        if "Tempo manual estimado" in self.cartoes_estatisticas:
            rotulo_valor, rotulo_extra = self.cartoes_estatisticas["Tempo manual estimado"]
            rotulo_valor.config(text=formatar_duracao(totais["duracao_manual_segundos"]))
            rotulo_extra.config(text="se feito à mão, com as estimativas acima")
        if "Tempo economizado" in self.cartoes_estatisticas:
            rotulo_valor, rotulo_extra = self.cartoes_estatisticas["Tempo economizado"]
            estilo = "NumeroGrandeOk.TLabel" if totais["tempo_economizado_segundos"] >= 0 else "Erro.TLabel"
            _configurar(
                rotulo_valor, text=formatar_duracao(abs(totais["tempo_economizado_segundos"])), style=estilo,
            )
            rotulo_extra.config(text=f"{texto_percentual(totais['percentual_economia'])} frente ao manual")

        for widget in self.frame_linhas_estatisticas.winfo_children():
            widget.destroy()
        for linha, categoria in enumerate(CATEGORIAS_ESTATISTICAS):
            bloco = dados["por_categoria"][categoria]
            valores = [
                categoria,
                str(bloco["processados"]),
                formatar_duracao(bloco["duracao_automacao_segundos"]),
                formatar_duracao(bloco["duracao_manual_segundos"]),
                formatar_duracao(abs(bloco["tempo_economizado_segundos"])),
                texto_percentual(bloco["percentual_economia"]) if bloco["processados"] else "-",
            ]
            for coluna, valor in enumerate(valores):
                ttk.Label(
                    self.frame_linhas_estatisticas, text=valor, style="Instrucao.TLabel", width=15,
                ).grid(row=linha, column=coluna, sticky="w", padx=4, pady=2)

        for widget in self.frame_linhas_projecao.winfo_children():
            widget.destroy()
        for linha, bloco in enumerate(dados["projecao"]):
            valores = [
                f"{bloco['itens']:,}".replace(",", "."),
                formatar_duracao(bloco["tempo_automatizado_segundos"]),
                formatar_duracao(bloco["tempo_manual_segundos"]),
                formatar_duracao(bloco["tempo_economizado_segundos"]),
            ]
            for coluna, valor in enumerate(valores):
                ttk.Label(
                    self.frame_linhas_projecao, text=valor, style="Instrucao.TLabel", width=20,
                ).grid(row=linha, column=coluna, sticky="w", padx=4, pady=3)
        _ajustar_cartoes(self.frame_linhas_estatisticas)
        _ajustar_cartoes(self.frame_linhas_projecao)

    # ------------------------------------------------------------------
    # ABA: RELATÓRIO
    #
    # Mostra quantas CE/NL/PP/OB foram feitas num mês escolhido pelo
    # usuário e quanto tempo cada uma levou, a partir do histórico salvo
    # em `amigo.relatorio` (historico.json) - 1 registro por automação
    # executada, gravado automaticamente pela aba "Automações"
    # (`_executar_automacao`), com sucesso ou com erro. Os dados ficam
    # guardados no disco, então continuam disponíveis mesmo depois de
    # fechar e reabrir o programa.
    # ------------------------------------------------------------------
    def _montar_aba_relatorio(self):
        conteudo = _criar_area_rolavel(self.aba_relatorio)

        _criar_cabecalho_aba(
            conteudo,
            "📅  Relatório do mês",
            "Quantas CE, NL, PP e OB foram feitas no mês escolhido e quanto tempo levaram. "
            "Fica guardado neste computador, mesmo depois de fechar o programa.",
        )

        topo = ttk.Frame(conteudo)
        topo.pack(fill="x", pady=(0, 22))
        ttk.Label(topo, text="📆  Mês:", style="Subtitulo.TLabel").pack(side="left", padx=(0, 10))
        self.var_mes_relatorio = tk.StringVar()
        self.combo_mes_relatorio = ttk.Combobox(
            topo, textvariable=self.var_mes_relatorio, state="readonly", width=10,
        )
        self.combo_mes_relatorio.pack(side="left")
        self.combo_mes_relatorio.bind("<<ComboboxSelected>>", lambda _evento: self._atualizar_relatorio())
        ttk.Button(topo, text="🔄  Atualizar", command=self._atualizar_relatorio).pack(side="left", padx=(10, 0))

        # ---- Resumo do mês: 1 "cartão" por categoria (CE / NL / PP / OB) ----
        frame_resumo = ttk.Frame(conteudo)
        frame_resumo.pack(fill="x", pady=(0, 6), padx=2)
        self.cartoes_relatorio = {}
        for indice, categoria in enumerate(CATEGORIAS_PRINCIPAIS):
            cartao = Cartao(frame_resumo, text=f"📄  {categoria}", padding=12)
            cartao.grid(row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 16, 0))
            frame_resumo.columnconfigure(indice, weight=1)
            rotulo_quantidade = ttk.Label(cartao, text="0 feita(s)", style="Subtitulo.TLabel")
            rotulo_quantidade.pack(anchor="w")
            rotulo_tempo = ttk.Label(cartao, text="Tempo total: -", style="Instrucao.TLabel")
            rotulo_tempo.pack(anchor="w", pady=(6, 0))
            rotulo_media = ttk.Label(cartao, text="Média por item: -", style="Instrucao.TLabel")
            rotulo_media.pack(anchor="w")
            self.cartoes_relatorio[categoria] = (rotulo_quantidade, rotulo_tempo, rotulo_media)

        self.rotulo_resumo_relatorio = ttk.Label(conteudo, text="", style="Instrucao.TLabel")
        self.rotulo_resumo_relatorio.pack(anchor="w", pady=(10, 18), padx=2)

        # ---- Detalhamento: 1 linha por automação (CE, Raspar CE, NL...) ----
        frame_detalhe = Cartao(conteudo, text="🧾  Detalhe de cada automação", padding=12)
        frame_detalhe.pack(fill="x", pady=(0, 22))

        cabecalho = ttk.Frame(frame_detalhe)
        cabecalho.pack(fill="x")
        titulos_colunas = ["Automação", "Execuções", "Itens processados", "Tempo total", "Média por item"]
        for coluna, titulo in enumerate(titulos_colunas):
            ttk.Label(
                cabecalho, text=titulo, style="Campo.TLabel", width=18 if coluna == 0 else 16,
            ).grid(row=0, column=coluna, sticky="w", padx=4, pady=(0, 8))

        self.frame_linhas_relatorio = ttk.Frame(frame_detalhe)
        self.frame_linhas_relatorio.pack(fill="x")

        self._atualizar_relatorio()

    def _popular_meses_relatorio(self):
        """Preenche o seletor de mês com os meses que já têm alguma
        execução registrada (mais o mês atual), preservando a seleção do
        usuário quando ela continua existindo na lista."""
        selecao_atual = self.var_mes_relatorio.get()
        meses = meses_disponiveis()
        self.combo_mes_relatorio["values"] = meses
        if selecao_atual in meses:
            self.var_mes_relatorio.set(selecao_atual)
        elif meses:
            self.var_mes_relatorio.set(meses[0])

    def _atualizar_relatorio(self):
        """Recarrega o resumo do mês escolhido a partir do histórico em
        disco. Chamada ao abrir a aba, ao trocar de mês, no botão
        'Atualizar' e automaticamente assim que qualquer automação
        termina (ver `_executar_automacao`)."""
        if not hasattr(self, "var_mes_relatorio"):
            return  # aba ainda não foi montada

        self._popular_meses_relatorio()
        mes_ano = self.var_mes_relatorio.get() or datetime.now().strftime("%m/%Y")
        resumo = resumo_mes(mes_ano)

        for categoria, (rotulo_qtd, rotulo_tempo, rotulo_media) in self.cartoes_relatorio.items():
            bloco = resumo["categorias"].get(categoria, {})
            processados = bloco.get("processados", 0)
            rotulo_qtd.config(text=f"{processados} feita(s)")
            rotulo_tempo.config(
                text=f"Tempo total: {formatar_duracao(bloco.get('duracao_total_segundos', 0))}"
            )
            rotulo_media.config(
                text=(
                    f"Média por item: {formatar_duracao(bloco.get('duracao_media_segundos', 0))}"
                    if processados else "Média por item: -"
                )
            )

        if resumo["total_execucoes"]:
            self.rotulo_resumo_relatorio.config(
                text=f"{resumo['total_execucoes']} automação(ões) executada(s) em {resumo['mes_ano']} "
                     "no total (inclui as etapas de conferência/raspagem, detalhadas abaixo)."
            )
        else:
            self.rotulo_resumo_relatorio.config(
                text=f"Nenhuma automação foi executada em {resumo['mes_ano']} ainda."
            )

        for widget in self.frame_linhas_relatorio.winfo_children():
            widget.destroy()

        detalhamento_por_programa = {item["programa"]: item for item in resumo["detalhamento"]}
        nomes_programas = [nome for _numero, (nome, _funcao) in sorted(PROGRAMAS.items())]
        bloco_vazio = {"execucoes": 0, "processados": 0, "duracao_total_segundos": 0, "duracao_media_segundos": 0}
        for linha, nome_programa in enumerate(nomes_programas):
            bloco = detalhamento_por_programa.get(nome_programa, bloco_vazio)
            valores = [
                nome_programa,
                str(bloco["execucoes"]),
                str(bloco["processados"]),
                formatar_duracao(bloco["duracao_total_segundos"]),
                formatar_duracao(bloco["duracao_media_segundos"]) if bloco["processados"] else "-",
            ]
            for coluna, valor in enumerate(valores):
                ttk.Label(
                    self.frame_linhas_relatorio, text=valor, style="Instrucao.TLabel",
                    width=18 if coluna == 0 else 16,
                ).grid(row=linha, column=coluna, sticky="w", padx=4, pady=3)
        _ajustar_cartoes(self.frame_linhas_relatorio)


def _ativar_nitidez_windows() -> None:
    """
    Avisa o Windows que o programa sabe desenhar na resolução real da tela.
    Sem isso, com a tela em 125%/150% (comum em notebooks), o Windows
    desenha o programa pequeno e depois "estica" a imagem - e tudo fica
    borrado/embaçado assim que a janela abre. Precisa rodar ANTES de criar
    a janela. Em outros sistemas não faz nada.
    """
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # nítido na resolução do monitor principal
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # Windows mais antigo
        except Exception:
            pass


def iniciar():
    """Ponto de entrada da interface gráfica (chamado por main_gui.py)."""
    _ativar_nitidez_windows()
    root = tk.Tk()
    _aplicar_tema_escuro(root)
    AplicativoAmigo(root)
    root.mainloop()
