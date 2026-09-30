# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
CONTROLE DE EXECUÇÃO DAS AUTOMAÇÕES
===================================
Três coisas que valem para QUALQUER automação que esteja rodando, num só
lugar (só roda uma automação por vez, então um estado único basta):

  - MODO SIMULAÇÃO: a automação faz tudo no SIGEF de verdade (pesquisa o
    credor, o empenho, a conta bancária, confere os valores...), mas PARA
    ANTES do clique final que confirmaria a operação e limpa a tela. Nada
    é gerado no SIGEF. No lugar do número (CE/NL/PP/OB) fica
    "SIMULADO ✓ ..." (tudo certo) ou "SIMULADO ✗ motivo" (o que daria
    errado). Como o SIGEF só cria o número DEPOIS da confirmação, não há
    número para coletar numa simulação.

  - PARAR: o botão "Parar" pede a parada; a automação termina a linha que
    está fazendo e para antes da próxima (nunca no meio de uma linha, para
    não deixar um lançamento pela metade).

  - PROGRESSO: cada automação avisa "linha X de Y"; a interface mostra
    isso na barra de andamento.

Este módulo não conhece Tkinter: a interface só registra ouvintes e chama
as funções daqui.
"""
import re
import threading
from contextlib import contextmanager

from .log import adicionar_ouvinte, log_aviso, log_sucesso, remover_ouvinte

MARCA_SIMULADO = "SIMULADO"

_parada = threading.Event()
_estado = {"simulacao": False}
_ouvintes_progresso = []


# ---- início / fim ----------------------------------------------------------
def iniciar_execucao(simulacao: bool = False) -> None:
    """Chamada ANTES de rodar uma automação: zera o pedido de parada e liga
    (ou desliga) o modo simulação."""
    _parada.clear()
    _estado["simulacao"] = bool(simulacao)


def finalizar_execucao() -> None:
    """Chamada DEPOIS da automação: volta ao normal (simulação desligada)."""
    _parada.clear()
    _estado["simulacao"] = False


# ---- simulação -------------------------------------------------------------
def em_simulacao() -> bool:
    return _estado["simulacao"]


def eh_simulado(valor) -> bool:
    """A célula traz um resultado de simulação (e não um número de verdade)?
    As automações usam isto para NÃO tentar pesquisar "SIMULADO ✓..." no
    SIGEF como se fosse uma CE/NL/PP."""
    return str(valor or "").strip().upper().startswith(MARCA_SIMULADO)


def texto_simulado_ok(etapa: str, detalhe: str = "") -> str:
    texto = f"{MARCA_SIMULADO} ✓ {etapa} conferida - nada foi gerado"
    return f"{texto} ({detalhe})" if detalhe else texto


def texto_simulado_problema(motivo: str) -> str:
    return f"{MARCA_SIMULADO} ✗ {motivo}"


# ---- parada ----------------------------------------------------------------
def pedir_parada() -> None:
    _parada.set()


def parada_pedida() -> bool:
    return _parada.is_set()


def parar_antes_da_linha(numero_linha) -> bool:
    """Chamada no começo de cada linha: True se o usuário pediu para parar
    (e registra no painel onde a automação parou)."""
    if _parada.is_set():
        log_aviso(
            f"Parada pedida pelo usuário: a automação parou antes da linha {numero_linha}. "
            f"Esta linha e as seguintes não foram feitas."
        )
        return True
    return False


# ---- progresso -------------------------------------------------------------
def adicionar_ouvinte_progresso(ouvinte) -> None:
    """`ouvinte(atual, total)` - chamado a cada linha iniciada."""
    if ouvinte not in _ouvintes_progresso:
        _ouvintes_progresso.append(ouvinte)


def remover_ouvinte_progresso(ouvinte) -> None:
    if ouvinte in _ouvintes_progresso:
        _ouvintes_progresso.remove(ouvinte)


def informar_progresso(atual: int, total: int) -> None:
    for ouvinte in list(_ouvintes_progresso):
        try:
            ouvinte(atual, total)
        except Exception:
            pass  # a tela nunca pode derrubar a automação


# ---- 1 linha de uma automação que GERA documento (CE, NL, PP) ----------------
_PREFIXO_LINHA = re.compile(r"^Linha \d+:\s*")


class LinhaEmExecucao:
    """Acompanha 1 linha: se, na simulação, a linha foi TENTADA mas não
    chegou ao ponto de confirmar, grava o motivo (a última mensagem de
    aviso/erro da automação) na coluna do documento."""

    def __init__(self, etapa, worksheet, numero_linha, coluna):
        self.etapa = etapa
        self.worksheet = worksheet
        self.numero_linha = numero_linha
        self.coluna = coluna
        self.tentou = False
        self.ok = False
        self.mensagens = []

    def tentando(self) -> None:
        """Os dados da linha estão preenchidos e a automação vai tentar."""
        self.tentou = True

    def simulado_ok(self, detalhe: str = "") -> None:
        """Chegou ao ponto de confirmar - na simulação, para aqui."""
        from .excel import salvar_valor_gerado

        self.ok = True
        log_sucesso(
            f"Linha {self.numero_linha}: SIMULAÇÃO - {self.etapa} conferida e pronta para "
            f"confirmar. Parou antes do clique final: nada foi gerado no SIGEF."
        )
        salvar_valor_gerado(
            self.worksheet, self.numero_linha, self.coluna,
            texto_simulado_ok(self.etapa, detalhe), rotulo=f"{self.etapa} (simulação)",
        )


@contextmanager
def linha_em_execucao(etapa, worksheet, numero_linha, coluna, atual, total):
    """
    Envolve o processamento de 1 linha das automações que geram documento:

        with linha_em_execucao("CE", worksheet, numero_linha, COL, i + 1, len(dados)) as ctx:
            ...                      # valida os dados da linha
            ctx.tentando()
            ...                      # preenche a tela
            if em_simulacao():
                ctx.simulado_ok(...)  # para ANTES de confirmar
                continue
            ...                      # confirma (só no modo normal)

    Fora da simulação, só informa o progresso - o comportamento da
    automação fica exatamente o mesmo de antes.
    """
    informar_progresso(atual, total)
    contexto = LinhaEmExecucao(etapa, worksheet, numero_linha, coluna)
    ouvinte = None
    if em_simulacao():
        def ouvinte(nivel, mensagem):
            if nivel in ("aviso", "erro"):
                contexto.mensagens.append(mensagem)
        adicionar_ouvinte(ouvinte)
    try:
        yield contexto
    finally:
        if ouvinte is not None:
            remover_ouvinte(ouvinte)
        if em_simulacao() and contexto.tentou and not contexto.ok:
            from .excel import atualizar_status

            motivo = contexto.mensagens[-1] if contexto.mensagens else f"a {etapa} não pôde ser conferida"
            motivo = _PREFIXO_LINHA.sub("", motivo)
            atualizar_status(worksheet, numero_linha, coluna, texto_simulado_problema(motivo))
