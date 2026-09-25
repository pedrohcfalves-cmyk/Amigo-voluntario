"""
COLUNAS RETRÁTEIS
=================
Tudo relacionado a "em qual coluna da planilha está cada dado". Define:

  - O layout de fábrica (as mesmas colunas que o sistema sempre usou).
  - `aplicar_colunas()`: recalcula as constantes COL_* a partir da
    configuração do usuário (config.json -> chave "colunas").
  - O submenu interativo "Configurar colunas" (tabela visual, bloqueio de
    colunas duplicadas, restaurar padrão).

As automações (pacote `amigo.automacoes`) NUNCA usam um número de coluna
fixo: sempre leem `colunas.COL_*`, então basta reatribuir esses globais
aqui (via `aplicar_colunas`) para todo o sistema passar a usar o novo
layout - sem tocar em nenhuma automação.
"""
from typing import Optional

from .config_io import salvar_configuracoes
from .log import log_info, log_sucesso, log_aviso

def indice_para_letra_coluna(indice: int) -> str:
    """
    Converte um índice de coluna base 1 (1, 2, 3, ... 13, ...) para a
    letra correspondente no Excel (A, B, C, ... M, ...).
    """
    letras = ""
    while indice > 0:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def letra_para_indice_coluna(letra: str) -> int:
    """
    Converte uma letra de coluna do Excel (A, B, ... Z, AA, AB, ...) para o
    índice base 1 correspondente (inverso de `indice_para_letra_coluna`).
    Usada por `aplicar_colunas()` para transformar as letras escolhidas
    pelo usuário no submenu "Configurar colunas" nos números de coluna que
    o restante do sistema usa internamente.

    Lança `ValueError` se `letra` não for uma sequência de letras válida.
    """
    letra = (letra or "").strip().upper()
    if not letra or not letra.isalpha():
        raise ValueError(f"Letra de coluna inválida: '{letra}'")
    indice = 0
    for caractere in letra:
        indice = indice * 26 + (ord(caractere) - ord("A") + 1)
    return indice



# Coluna mínima (base 1) usada por `obter_proxima_coluna_livre()` para
# localizar a primeira coluna livre de uma linha, caso seja necessária.
COL_COLUNA_MINIMA = 8    # Coluna H

# ---- Colunas específicas da automação CE (base 0, usadas com `obter_celula`) --
COL_CE_CPF = 1      # Coluna B - CPF do credor
COL_CE_VALOR = 7    # Coluna H - valor do documento (será convertido para centavos)

# Coluna (base 1) onde o número da CE gerada é salvo, na mesma linha.
COL_CE_GERADA = 9   # Coluna I

# ---- Colunas específicas da automação NL (base 0, usadas com `obter_celula`) --
COL_NL_CE = 8       # Coluna I - número da CE já gerada (mesma coluna gravada pela CE/Raspar CE)
COL_NL_NE = 3       # Coluna D - número da Nota de Empenho vinculada

# Coluna (base 1) onde o número da NL gerada é salvo.
COL_NL_GERADA = 10  # Coluna J

# ---- Colunas específicas da automação PP (base 0, usadas com `obter_celula`) --
# Mapeamento definido junto com a automação PP - não alterar o layout da
# planilha nem o posicionamento destas colunas.
COL_PP_BANCO = 4    # Coluna E - Banco do domicílio bancário (ex: 1)
COL_PP_AGENCIA = 5  # Coluna F - Agência bancária (ex: 1178-9)
COL_PP_CONTA = 6    # Coluna G - Conta bancária (ex: 74.508-1)
COL_PP_CE = COL_NL_CE  # Coluna I - Despesa Certificada (mesma coluna lida pela NL)
COL_PP_NL = 9       # Coluna J - Nota de Lançamento gerada pela NL (COL_NL_GERADA em base 0)

# Coluna (base 1) onde o número da PP gerada é salvo. Quando o SIGEF recusa
# o lançamento (ex: "Saldo insuficiente..."), a mensagem devolvida por ele é
# gravada nesta MESMA coluna, para que a planilha mostre o motivo da falha.
COL_PP_GERADA = 11  # Coluna K

# ---- Colunas da automação Raspar PP -------------------------------------
# Entrada (base 0, usadas com `obter_celula`):
COL_RASPAR_PP_CPF = 1              # Coluna B - CPF do favorecido
COL_RASPAR_PP_VALOR = COL_CE_VALOR # Coluna H - valor usado para achar a PP na grade
COL_RASPAR_PP_CE = COL_NL_CE       # Coluna I - CE esperada
COL_RASPAR_PP_NL = COL_PP_NL       # Coluna J - NL esperada
# Saída: a PP raspada é gravada na MESMA coluna K usada pela automação PP
# (COL_PP_GERADA) - de propósito, para que as duas nunca gravem em lugares
# diferentes.

# Coluna (base 1) da mensagem de conferência da Raspar PP.
# ATENÇÃO - COLISÃO: as duas constantes abaixo apontam para a coluna M:
#   COL_RASPAR_CONTA_RESULTADO  (Raspar contas)
#   COL_RASPAR_PP_OBSERVACAO    (Raspar PP)
# Rodar Raspar PP e depois Raspar contas SOBRESCREVE a mensagem da primeira.
# Se as duas precisarem coexistir na planilha, mova uma delas para outra
# coluna - basta alterar o número aqui, nada mais depende disso.
COL_RASPAR_PP_OBSERVACAO = 13      # Coluna M

# ---- Colunas da automação Raspar contas ---------------------------------
# Entrada (base 0, usadas com `obter_celula`):
# O CPF não é mais usado: a pesquisa passou a ser feita direto pelo número
# da PP (coluna K), sem o popup de favorecido.
COL_RASPAR_CONTA_VALOR = COL_CE_VALOR  # Coluna H - valor, usado para escolher
                                       # a linha certa quando a grade traz
                                       # mais de uma
COL_RASPAR_CONTA_PP = 10   # Coluna K  - PP já gerada (COL_PP_GERADA em base 0)
# O banco/agência/conta esperados são os MESMOS lidos pela automação PP
# (colunas E, F e G) - reaproveitados de propósito, para que as duas
# automações nunca comparem contra colunas diferentes.
COL_RASPAR_CONTA_BANCO = COL_PP_BANCO      # Coluna E
COL_RASPAR_CONTA_AGENCIA = COL_PP_AGENCIA  # Coluna F
COL_RASPAR_CONTA_CONTA = COL_PP_CONTA      # Coluna G

# Saída (base 1, usadas com `worksheet.Cells`):
# ATENÇÃO: a coluna M é a MESMA de COL_RASPAR_PP_OBSERVACAO ("Raspar PP") -
# ver o aviso acima.
COL_RASPAR_CONTA_RESULTADO = 13    # Coluna M - banco/agência/conta lidos do SIGEF
COL_RASPAR_CONTA_CONFERENCIA = 14  # Coluna N - "Igual" ou "Diferente"

# ---- Colunas da automação "Gerar OB" (lote de 30 em 30) -----------------
# Entrada (base 0, usadas com `obter_celula`): reaproveita as MESMAS
# colunas já lidas pela CE (valor) e pela Raspar contas (PP) - de
# propósito, para nunca comparar contra colunas diferentes das outras
# automações.
COL_OB_CPF = COL_CE_CPF                    # Coluna B - CPF do favorecido
COL_OB_VALOR = COL_CE_VALOR                # Coluna H - "Valor a Receber"
COL_OB_PP_ESPERADO = COL_RASPAR_CONTA_PP   # Coluna K - PP já gerada

# Saída (base 1, usada com `worksheet.Cells`): nº da OB gerada por este
# lote OU, se o item não foi localizado na grade, a mensagem explicando
# o motivo (mesma coluna para os dois casos, igual à COL_PP_GERADA).
COL_OB_GERADA = 12  # Coluna L


# ==============================================================================
# COLUNAS RETRÁTEIS (configuráveis pelo usuário no menu)
# ==============================================================================
# Tudo acima é o layout PADRÃO DE FÁBRICA (o mesmo que o sistema sempre
# usou). A partir daqui esse layout passa a poder ser trocado pelo usuário,
# pelo submenu "Configurar colunas" do menu principal, SEM tocar em nenhuma
# automação: elas continuam usando as MESMAS constantes COL_* de sempre
# (COL_CE_CPF, COL_PP_BANCO, COL_NL_GERADA etc.) - só que essas constantes
# passam a ser recalculadas por `aplicar_colunas()`, a partir da
# configuração salva em config.json (chave "colunas"), em vez de fixas.
#
# Cada item abaixo é (chave em config.json, rótulo exibido no menu, coluna
# padrão de fábrica - a mesma já usada hoje). O padrão só é usado a) na 1ª
# vez que o sistema roda (config.json ainda não existe) e b) se o usuário
# escolher "Restaurar colunas para o padrão de fábrica".

CAMPOS_COLUNAS_BASICAS = [
    # Colunas fundamentais: TODAS as automações dependem delas para saber
    # de quem é o pagamento e por quanto - por isso são obrigatórias e
    # reaproveitadas por CE, NL, PP, Raspar PP, Raspar contas e Gerar OB.
    ("cpf", "CPF do credor/favorecido", "B"),
    ("nota_empenho", "Nota de Empenho (NE)", "D"),
    ("banco", "Banco", "E"),
    ("agencia", "Agência", "F"),
    ("conta", "Conta", "G"),
    ("valor", "Valor", "H"),
]

CAMPOS_COLUNAS_ENCADEADAS = [
    # Colunas dos documentos gerados pelo SIGEF. Cada uma é, ao mesmo
    # tempo, ONDE a automação correspondente SALVA o número gerado e ONDE a
    # automação seguinte da cadeia (CE -> NL -> PP -> OB) vai BUSCAR esse
    # mesmo número. Ex: mudar "ce" para a coluna F faz a CE ser salva em F
    # E faz a NL (que depende da CE) passar a ler a CE também em F.
    ("ce", "CE - Despesa Certificada (gerada pela CE/Raspar CE; lida pela NL/Raspar PP)", "I"),
    ("nl", "NL - Nota de Lançamento (gerada pela NL/Raspar NL; lida pela PP/Raspar PP)", "J"),
    ("pp", "PP - Preparação de Pagamento (gerada pela PP/Raspar PP; lida pela Raspar contas/Gerar OB)", "K"),
    ("ob", "OB - Ordem Bancária (gerada pela Gerar OB)", "L"),
]

CAMPOS_COLUNAS_AVANCADAS = [
    # Colunas auxiliares (mensagens/conferências). Não fazem parte da
    # cadeia CE -> NL -> PP -> OB, mas também podem ser reposicionadas caso
    # colidam com alguma coluna que o usuário já usa na própria planilha.
    ("raspar_pp_observacao", "Observação da Raspar PP", "M"),
    ("raspar_conta_resultado", "Resultado (banco/agência/conta) da Raspar contas", "M"),
    ("raspar_conta_conferencia", "Conferência ('Igual'/'Diferente') da Raspar contas", "N"),
]

TODOS_OS_CAMPOS_COLUNAS = CAMPOS_COLUNAS_BASICAS + CAMPOS_COLUNAS_ENCADEADAS + CAMPOS_COLUNAS_AVANCADAS

# Fonte única do padrão de fábrica (chave -> letra), derivada dos campos
# acima. Usada por `aplicar_colunas()` e reaproveitada por
# `amigo.config.CONFIG_PADRAO["colunas"]` (que faz `dict(COLUNAS_PADRAO)`),
# para nunca existirem dois lugares com o layout padrão escrito por
# extenso (e o risco de um dia ficarem diferentes).
COLUNAS_PADRAO = {chave: padrao for chave, _rotulo, padrao in TODOS_OS_CAMPOS_COLUNAS}

# Coluna (base 1) até onde `ler_dados()` precisa ler a planilha para que
# `obter_celula()` enxergue TODAS as colunas configuradas. Recalculada por
# `aplicar_colunas()` sempre que a configuração muda; o valor abaixo (12=L)
# é só o ponto de partida, igual ao intervalo fixo "A:L" de sempre.
COL_ULTIMA_LEITURA = 12  # Coluna L


def aplicar_colunas(colunas: dict) -> None:
    """
    Recalcula TODAS as constantes COL_* (usadas por CE, NL, PP, Raspar PP,
    Raspar contas e Gerar OB) a partir do dicionário `colunas` - as letras
    configuradas pelo usuário, uma por campo lógico (ver
    `TODOS_OS_CAMPOS_COLUNAS`).

    Chamada automaticamente sempre que a configuração é carregada
    (`carregar_configuracoes()`) ou alterada pelo submenu "Configurar
    colunas". Nenhuma automação precisa ser alterada: elas continuam lendo
    as mesmas constantes globais de sempre, só que com o valor configurado
    pelo usuário em vez do valor fixo original.
    """
    global COL_CE_CPF, COL_CE_VALOR, COL_CE_GERADA
    global COL_NL_CE, COL_NL_NE, COL_NL_GERADA
    global COL_PP_BANCO, COL_PP_AGENCIA, COL_PP_CONTA, COL_PP_CE, COL_PP_NL, COL_PP_GERADA
    global COL_RASPAR_PP_CPF, COL_RASPAR_PP_VALOR, COL_RASPAR_PP_CE, COL_RASPAR_PP_NL
    global COL_RASPAR_PP_OBSERVACAO
    global COL_RASPAR_CONTA_VALOR, COL_RASPAR_CONTA_PP, COL_RASPAR_CONTA_BANCO
    global COL_RASPAR_CONTA_AGENCIA, COL_RASPAR_CONTA_CONTA
    global COL_RASPAR_CONTA_RESULTADO, COL_RASPAR_CONTA_CONFERENCIA
    global COL_OB_CPF, COL_OB_VALOR, COL_OB_PP_ESPERADO, COL_OB_GERADA
    global COL_COLUNA_MINIMA, COL_ULTIMA_LEITURA

    colunas = colunas or {}
    padroes = COLUNAS_PADRAO

    def indice_base1(chave: str) -> int:
        padrao = padroes[chave]
        letra_configurada = str(colunas.get(chave) or padrao).strip()
        try:
            return letra_para_indice_coluna(letra_configurada)
        except ValueError:
            log_aviso(
                f"Coluna configurada inválida para '{chave}' ('{letra_configurada}'). "
                f"Usando o padrão de fábrica '{padrao}'."
            )
            return letra_para_indice_coluna(padrao)

    cpf1 = indice_base1("cpf")
    nota_empenho1 = indice_base1("nota_empenho")
    banco1 = indice_base1("banco")
    agencia1 = indice_base1("agencia")
    conta1 = indice_base1("conta")
    valor1 = indice_base1("valor")
    ce1 = indice_base1("ce")
    nl1 = indice_base1("nl")
    pp1 = indice_base1("pp")
    ob1 = indice_base1("ob")
    raspar_pp_observacao1 = indice_base1("raspar_pp_observacao")
    raspar_conta_resultado1 = indice_base1("raspar_conta_resultado")
    raspar_conta_conferencia1 = indice_base1("raspar_conta_conferencia")

    # ---- Entrada (base 0 - lidas de `linha[indice]` via `obter_celula`) ----
    COL_CE_CPF = cpf1 - 1
    COL_CE_VALOR = valor1 - 1
    COL_NL_CE = ce1 - 1
    COL_NL_NE = nota_empenho1 - 1
    COL_PP_BANCO = banco1 - 1
    COL_PP_AGENCIA = agencia1 - 1
    COL_PP_CONTA = conta1 - 1
    COL_PP_CE = COL_NL_CE
    COL_PP_NL = nl1 - 1
    COL_RASPAR_PP_CPF = cpf1 - 1
    COL_RASPAR_PP_VALOR = COL_CE_VALOR
    COL_RASPAR_PP_CE = COL_NL_CE
    COL_RASPAR_PP_NL = COL_PP_NL
    COL_RASPAR_CONTA_VALOR = COL_CE_VALOR
    COL_RASPAR_CONTA_PP = pp1 - 1
    COL_RASPAR_CONTA_BANCO = COL_PP_BANCO
    COL_RASPAR_CONTA_AGENCIA = COL_PP_AGENCIA
    COL_RASPAR_CONTA_CONTA = COL_PP_CONTA
    COL_OB_CPF = COL_CE_CPF
    COL_OB_VALOR = COL_CE_VALOR
    COL_OB_PP_ESPERADO = COL_RASPAR_CONTA_PP

    # ---- Saída (base 1 - gravadas com `worksheet.Cells(linha, coluna)`) ----
    COL_CE_GERADA = ce1
    COL_NL_GERADA = nl1
    COL_PP_GERADA = pp1
    COL_OB_GERADA = ob1
    COL_RASPAR_PP_OBSERVACAO = raspar_pp_observacao1
    COL_RASPAR_CONTA_RESULTADO = raspar_conta_resultado1
    COL_RASPAR_CONTA_CONFERENCIA = raspar_conta_conferencia1

    # A "coluna mínima" para localizar uma coluna livre sempre começou no
    # mesmo lugar da coluna de Valor - continua assim, só que dinâmica.
    COL_COLUNA_MINIMA = valor1

    # `ler_dados()` precisa cobrir todas as colunas que alguma automação LÊ
    # de volta da lista `linha` (via `obter_celula`): as básicas e as
    # geradas que são reaproveitadas como entrada por outra automação (CE,
    # NL, PP). OB e as colunas "avançadas" nunca são lidas de volta - só
    # gravadas diretamente por `worksheet.Cells()` -, então não precisam
    # entrar aqui; isso mantém o intervalo padrão idêntico ao "A:L" fixo de
    # sempre, e só o amplia se o usuário mover um campo LIDO para além de L.
    COL_ULTIMA_LEITURA = max(
        cpf1, nota_empenho1, banco1, agencia1, conta1, valor1, ce1, nl1, pp1,
        12,  # nunca menor que o intervalo padrão original (A:L)
    )

    _avisar_colunas_duplicadas(colunas)


def _detectar_colisoes(colunas: dict) -> dict:
    """
    Verifica se dois (ou mais) campos lógicos diferentes acabam apontando
    para a MESMA coluna da planilha - o que faria uma automação ler ou
    gravar o dado errado. Retorna um dicionário {letra: [(chave, rótulo),
    ...]} só com as letras realmente em conflito.

    Exceção: o padrão de fábrica já usa a MESMA coluna M para
    'raspar_pp_observacao' e 'raspar_conta_resultado' (comportamento
    original do sistema, documentado nos comentários da seção "COLUNAS DA
    PLANILHA") - essa combinação específica, sozinha, não conta como
    conflito. Qualquer OUTRA coincidência (incluindo um terceiro campo
    colidindo com essas duas) conta normalmente.

    Usada tanto para BLOQUEAR o salvamento no submenu "Configurar colunas"
    (o usuário precisa corrigir antes de conseguir salvar) quanto para o
    aviso não-bloqueante de `_avisar_colunas_duplicadas()`.
    """
    chaves_avancadas = {chave for chave, _rotulo, _padrao in CAMPOS_COLUNAS_AVANCADAS}
    por_letra: dict = {}
    for chave, rotulo, padrao in TODOS_OS_CAMPOS_COLUNAS:
        letra = str(colunas.get(chave) or padrao).strip().upper()
        por_letra.setdefault(letra, []).append((chave, rotulo))

    colisoes = {}
    for letra, itens in por_letra.items():
        if len(itens) < 2:
            continue
        chaves_envolvidas = {chave for chave, _rotulo in itens}
        if chaves_envolvidas <= chaves_avancadas:
            continue
        colisoes[letra] = itens
    return colisoes


def _avisar_colunas_duplicadas(colunas: dict) -> None:
    """
    Versão não-bloqueante de `_detectar_colisoes()`. Usada por
    `aplicar_colunas()`, que também roda ao carregar o config.json direto
    do disco (ex: na inicialização do programa, sem ninguém no teclado
    para responder a uma pergunta) - por isso só registra um aviso no
    log, sem impedir o sistema de continuar. Quem impede é o submenu
    "Configurar colunas" (`_configurar_grupo_colunas`), que BLOQUEIA o
    salvamento enquanto houver conflito.
    """
    for letra, itens in _detectar_colisoes(colunas).items():
        rotulos = ", ".join(rotulo for _chave, rotulo in itens)
        log_aviso(
            f"A coluna {letra} está configurada para mais de um dado ao mesmo tempo "
            f"({rotulos}). Isso pode fazer a automação ler ou gravar o valor errado."
        )



def _formatar_coluna(letra: str) -> str:
    """Formata uma letra de coluna para exibição (maiúscula, ou aviso se vazia)."""
    return letra.strip().upper() if letra else "(vazio)"


# Nome curto de cada campo, usado só na TABELA (a pergunta em si continua
# usando o rótulo completo e descritivo de CAMPOS_COLUNAS_*).
NOMES_CURTOS_COLUNAS = {
    "cpf": "CPF",
    "nota_empenho": "Nota de Empenho",
    "banco": "Banco",
    "agencia": "Agência",
    "conta": "Conta",
    "valor": "Valor",
    "ce": "CE (Despesa Certificada)",
    "nl": "NL (Nota de Lançamento)",
    "pp": "PP (Preparação Pagamento)",
    "ob": "OB (Ordem Bancária)",
    "raspar_pp_observacao": "Obs. Raspar PP",
    "raspar_conta_resultado": "Resultado Raspar contas",
    "raspar_conta_conferencia": "Conferência Raspar contas",
}

GRUPOS_CAMPOS_COLUNAS = {
    **{chave: "Básica" for chave, _rotulo, _padrao in CAMPOS_COLUNAS_BASICAS},
    **{chave: "Gerada/encadeada" for chave, _rotulo, _padrao in CAMPOS_COLUNAS_ENCADEADAS},
    **{chave: "Avançada" for chave, _rotulo, _padrao in CAMPOS_COLUNAS_AVANCADAS},
}


def _tabela_colunas(colunas: dict) -> str:
    """
    Monta uma TABELA (texto, com bordas) com uma linha por campo,
    ORDENADA pela letra da coluna - assim, se dois campos usarem a mesma
    coluna, eles aparecem em linhas vizinhas com a mesma letra, deixando
    qualquer sobreposição visualmente óbvia. A coluna "Situação" marca
    "!! DUPLICADA !!" nas linhas em conflito (ver `_detectar_colisoes`).
    """
    colisoes = _detectar_colisoes(colunas)

    linhas = []
    for chave, _rotulo, padrao in TODOS_OS_CAMPOS_COLUNAS:
        letra = _formatar_coluna(colunas.get(chave) or padrao)
        nome = NOMES_CURTOS_COLUNAS.get(chave, chave)
        grupo = GRUPOS_CAMPOS_COLUNAS.get(chave, "")
        situacao = "!! DUPLICADA !!" if letra in colisoes else "OK"
        linhas.append((letra, nome, grupo, situacao))

    def chave_ordenacao(linha):
        letra = linha[0]
        try:
            return (letra_para_indice_coluna(letra), letra)
        except ValueError:
            return (9999, letra)

    linhas.sort(key=chave_ordenacao)

    cabecalho = ("Coluna", "Dado", "Grupo", "Situação")
    todas_linhas = [cabecalho] + linhas
    larguras = [
        max(len(str(linha[i])) for linha in todas_linhas)
        for i in range(4)
    ]

    def borda(esquerda, meio, direita):
        return esquerda + meio.join("─" * (largura + 2) for largura in larguras) + direita

    def formatar_linha(colunas_linha):
        celulas = [f" {str(valor).ljust(largura)} " for valor, largura in zip(colunas_linha, larguras)]
        return "│" + "│".join(celulas) + "│"

    saida = [borda("┌", "┬", "┐"), formatar_linha(cabecalho), borda("├", "┼", "┤")]
    for linha in linhas:
        saida.append(formatar_linha(linha))
    saida.append(borda("└", "┴", "┘"))
    return "\n".join(saida)


class _CancelarConfiguracaoColunas(Exception):
    """Sinaliza que o usuário digitou 0/cancelar durante a configuração de
    colunas - usada só para interromper o loop de perguntas sem salvar
    nada, nunca é deixada escapar para fora de `_configurar_grupo_colunas`."""


def _perguntar_coluna(rotulo: str, atual: str, padrao: str) -> Optional[str]:
    """
    Pergunta a letra da coluna para um único campo, mostrando o valor JÁ
    CONFIGURADO como padrão da pergunta (é isso que fica selecionado se o
    usuário só apertar ENTER). O padrão de fábrica só é mostrado à parte,
    como referência, quando for diferente do valor atual.

    Retorna a nova letra (maiúscula) ou None se nada precisar mudar
    (ENTER vazio, ou letra inválida - caso em que o valor atual é mantido).
    Digitar "0" ou "cancelar" interrompe TODA a configuração em andamento
    (levanta `_CancelarConfiguracaoColunas`, tratada por quem chamou).
    """
    atual_formatado = _formatar_coluna(atual)
    sufixo_padrao = "" if atual_formatado == padrao.upper() else f" (padrão de fábrica: {padrao})"
    entrada = input(f"  {rotulo} [{atual_formatado}]{sufixo_padrao} (0 cancela tudo): ").strip()
    if not entrada:
        return None
    if entrada.lower() in ("0", "cancelar"):
        raise _CancelarConfiguracaoColunas()
    try:
        letra_para_indice_coluna(entrada)
    except ValueError:
        log_aviso(f"Letra de coluna inválida: '{entrada}'. Mantendo '{atual_formatado}'.")
        return None
    return entrada.upper()


def _configurar_grupo_colunas(config: dict, titulo: str, campos: list) -> dict:
    """
    Percorre um grupo de campos (básicos, encadeados ou avançados),
    perguntando a coluna de cada um. NADA é gravado em config.json nem
    aplicado às automações enquanto o usuário não CONFIRMAR: as respostas
    ficam numa cópia de trabalho, uma tabela de pré-visualização é
    mostrada, e só então vem a pergunta de confirmação.

    Duas colunas nunca podem ficar apontando para o mesmo dado: se a
    combinação de respostas produzir uma coluna duplicada (ver
    `_detectar_colisoes`), o salvamento é BLOQUEADO e o usuário é levado
    de volta só aos campos em conflito, para corrigir antes de continuar.
    """
    colunas_original = config.get("colunas")
    if not isinstance(colunas_original, dict):
        colunas_original = dict(COLUNAS_PADRAO)

    colunas_proposta = dict(colunas_original)  # cópia de trabalho - nada é salvo ainda

    print(f"\n🧱 {titulo}")
    print("-" * 50)
    print("Tabela atual (coluna -> dado):")
    print(_tabela_colunas(colunas_proposta))
    print("\nENTER em qualquer campo mantém a coluna já configurada.")

    campos_pendentes = list(campos)
    primeira_rodada = True
    try:
        while campos_pendentes:
            if not primeira_rodada:
                print("\nInforme novamente apenas os campos em conflito:")
            for chave, rotulo, padrao in campos_pendentes:
                atual = colunas_proposta.get(chave) or padrao
                nova_letra = _perguntar_coluna(rotulo, atual, padrao)
                if nova_letra:
                    colunas_proposta[chave] = nova_letra

            colisoes = _detectar_colisoes(colunas_proposta)
            if not colisoes:
                break

            print("\n❌ Não é possível salvar: as colunas abaixo ficaram duplicadas.")
            chaves_em_conflito = set()
            for letra, itens in colisoes.items():
                rotulos = ", ".join(rotulo for _chave, rotulo in itens)
                print(f"   Coluna {letra}: {rotulos}")
                chaves_em_conflito.update(chave for chave, _rotulo in itens)

            # Só volta a perguntar os campos DESTE grupo que causaram o
            # conflito - se o conflito for só com um campo de outro grupo
            # (já salvo antes), não há nada aqui para reperguntar.
            campos_pendentes = [
                (chave, rotulo, padrao) for chave, rotulo, padrao in campos
                if chave in chaves_em_conflito
            ]
            if not campos_pendentes:
                print(
                    "   (o conflito é com um campo configurado em outro grupo de colunas - "
                    "ajuste-o lá, ou mude um dos valores acima, antes de continuar)"
                )
                log_info("Alteração cancelada. Nenhuma coluna foi salva.")
                return config
            primeira_rodada = False
    except _CancelarConfiguracaoColunas:
        log_info("Alteração cancelada pelo usuário. Nenhuma coluna foi salva.")
        return config

    if colunas_proposta == colunas_original:
        log_info("Nenhuma alteração feita.")
        return config

    print("\nPré-visualização da nova configuração:")
    print(_tabela_colunas(colunas_proposta))
    confirmacao = input("\nConfirma gravar essas colunas? (s/N): ").strip().lower()
    if confirmacao != "s":
        log_info("Alteração cancelada. Nenhuma coluna foi salva.")
        return config

    config["colunas"] = colunas_proposta
    salvar_configuracoes(config)
    aplicar_colunas(colunas_proposta)
    log_sucesso("Colunas atualizadas com sucesso!")
    return config


def exibir_colunas_atuais(config: dict):
    """Exibe, em forma de tabela, em qual coluna cada dado está
    configurado hoje (a letra da coluna, exatamente como no Excel)."""
    colunas = config.get("colunas")
    if not isinstance(colunas, dict):
        colunas = dict(COLUNAS_PADRAO)

    print("\n📋 COLUNAS CONFIGURADAS ATUALMENTE")
    print(_tabela_colunas(colunas))


def restaurar_colunas_padrao(config: dict) -> dict:
    """Restaura todas as colunas para o padrão de fábrica (o layout que o
    sistema sempre usou: B, D, E, F, G, H, I, J, K, L, M, M, N)."""
    print("\n♻️  RESTAURAR COLUNAS PARA O PADRÃO DE FÁBRICA")
    print(_tabela_colunas(COLUNAS_PADRAO))
    confirmacao = input(
        "\nIsso volta TODAS as colunas para a tabela acima. Confirma? (s/N): "
    ).strip().lower()
    if confirmacao == "s":
        config["colunas"] = dict(COLUNAS_PADRAO)
        salvar_configuracoes(config)
        aplicar_colunas(config["colunas"])
        log_sucesso("Colunas restauradas para o padrão de fábrica.")
    else:
        log_info("Operação cancelada. Nenhuma alteração foi feita.")
    return config


def menu_configurar_colunas(config: dict) -> dict:
    """
    Submenu 'Colunas Retráteis': permite escolher em qual coluna cada dado
    deve ser lido ou salvo, para que o sistema funcione com o layout de
    QUALQUER planilha - sem alterar nenhuma automação nem o salvamento.

    O valor mostrado como padrão para cada campo é sempre o que já está
    configurado HOJE (indicado pela letra da coluna); da primeira vez que
    o sistema roda, esse valor é o layout de fábrica de sempre (B, D, E,
    F, G, H, I, J, K, L). Nenhuma coluna é salva duplicada: se a
    combinação de respostas colidir, o sistema pede para corrigir antes
    de continuar (ver `_configurar_grupo_colunas`).
    """
    while True:
        colunas = config.get("colunas")
        if not isinstance(colunas, dict):
            colunas = dict(COLUNAS_PADRAO)

        print("\n🧱 CONFIGURAR COLUNAS (Colunas Retráteis)")
        print("=" * 50)
        print(_tabela_colunas(colunas))
        print("=" * 50)
        print("1 - Colunas básicas (CPF, Nota de Empenho, Banco, Agência, Conta, Valor)")
        print("2 - Colunas de documentos gerados (CE, NL, PP, OB)")
        print("3 - Colunas avançadas (observação/resultado/conferência)")
        print("4 - Exibir colunas configuradas atualmente")
        print("5 - Restaurar colunas para o padrão de fábrica")
        print("0 - Voltar ao menu principal")
        print("=" * 50)

        escolha = input("Escolha uma opção: ").strip()

        if escolha == "1":
            config = _configurar_grupo_colunas(config, "COLUNAS BÁSICAS", CAMPOS_COLUNAS_BASICAS)
        elif escolha == "2":
            config = _configurar_grupo_colunas(
                config, "COLUNAS DE DOCUMENTOS GERADOS (CE -> NL -> PP -> OB)", CAMPOS_COLUNAS_ENCADEADAS
            )
        elif escolha == "3":
            config = _configurar_grupo_colunas(config, "COLUNAS AVANÇADAS", CAMPOS_COLUNAS_AVANCADAS)
        elif escolha == "4":
            exibir_colunas_atuais(config)
        elif escolha == "5":
            config = restaurar_colunas_padrao(config)
        elif escolha == "0":
            return config
        else:
            log_aviso("Opção inválida. Tente novamente.")

