# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
RELATÓRIO
=========
Histórico de execuções das automações: quando rodaram, quantas linhas
processaram e quanto tempo levaram. Persistido em `historico.json` (na
mesma pasta de `config.json`) para alimentar a aba "Relatório" da
interface gráfica - permite ver quantas CE/NL/PP/OB foram feitas num mês
específico e o tempo gasto, mesmo depois de fechar e reabrir o programa.

Cada execução de uma automação (aba "Automações" da GUI) grava 1 registro
aqui - com sucesso ou com erro - assim o relatório reflete o uso real do
sistema, não só as automações 100% bem-sucedidas. Quem grava é
`gui.AplicativoAmigo._executar_automacao`; este módulo só cuida de ler,
somar e gravar - não conhece nada de Tkinter.
"""
import json
import os
from datetime import datetime
from typing import List, Optional

from .config_io import BASE_DIR
from .historico_git import guardar_historico_no_git
from .log import log_erro

HISTORICO_PATH = os.path.join(BASE_DIR, "historico.json")

# Em qual "categoria" do resumo (CE / NL / PP / OB) cada automação entra -
# ver amigo.automacoes.PROGRAMAS para a lista completa. CE/Raspar CE
# contam juntas como "CE" (as duas fazem uma CE existir na planilha - uma
# gerando na hora, a outra recuperando uma já existente no SIGEF), e o
# mesmo raciocínio vale para NL/PP/OB. "Raspar contas" é só conferência
# (não gera nenhum documento novo), por isso fica de fora das 4
# categorias principais - aparece só no detalhamento por automação.
CATEGORIA_POR_PROGRAMA = {
    "CE": "CE", "Raspar CE": "CE",
    "NL": "NL", "Raspar NL": "NL",
    "PP": "PP", "Raspar PP": "PP",
    "Gerar OB": "OB", "Raspar OB": "OB", "Confirmar OB": "OB",
    "Raspar contas": "Conferência",
}
CATEGORIAS_PRINCIPAIS = ["CE", "NL", "PP", "OB"]


def _carregar_historico() -> List[dict]:
    if not os.path.exists(HISTORICO_PATH):
        return []
    try:
        with open(HISTORICO_PATH, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
            return dados if isinstance(dados, list) else []
    except (json.JSONDecodeError, OSError) as erro:
        log_erro(f"Não foi possível ler historico.json ({erro}). O relatório começa vazio.")
        return []


def _salvar_historico(historico: List[dict]) -> None:
    try:
        with open(HISTORICO_PATH, "w", encoding="utf-8") as arquivo:
            json.dump(historico, arquivo, ensure_ascii=False, indent=2)
    except OSError as erro:
        log_erro(f"Falha ao salvar historico.json: {erro}")


def registrar_execucao(
    programa: str,
    inicio: datetime,
    fim: datetime,
    processados: int,
    total_linhas: int,
    sucesso: bool,
    detalhe: Optional[str] = None,
) -> None:
    """
    Adiciona 1 execução ao histórico e salva em historico.json na hora
    (append simples - nunca reescreve execuções antigas). Chamada 1 vez
    ao final de CADA automação disparada pela aba "Automações", com
    sucesso ou com erro (ver `gui._executar_automacao`) - uma exceção
    aqui dentro NUNCA pode derrubar a automação que acabou de rodar, por
    isso quem chama envolve isso num try/except à parte.
    """
    historico = _carregar_historico()
    historico.append({
        "programa": programa,
        "categoria": CATEGORIA_POR_PROGRAMA.get(programa, programa),
        "inicio": inicio.isoformat(timespec="seconds"),
        "fim": fim.isoformat(timespec="seconds"),
        "duracao_segundos": max(0, round((fim - inicio).total_seconds())),
        "processados": int(processados or 0),
        "total_linhas": int(total_linhas or 0),
        "sucesso": bool(sucesso),
        "detalhe": detalhe or "",
    })
    _salvar_historico(historico)
    # Versiona o historico.json no Git (em segundo plano; se não houver
    # Git ou repositório, não faz nada - ver amigo/historico_git.py).
    guardar_historico_no_git(
        HISTORICO_PATH,
        f"Histórico: {programa} em {inicio.strftime('%d/%m/%Y %H:%M')} "
        f"({int(processados or 0)} de {int(total_linhas or 0)} linha(s))",
    )


def historico_completo() -> List[dict]:
    """
    Devolve TODO o histórico de execuções, sem filtro de mês - usada por
    `amigo.estatisticas` (aba "Estatísticas"), que precisa olhar o
    histórico completo (ou um período escolhido pelo usuário) para calcular
    tempo economizado e a relação tempo × itens. É só um "abridor" público
    para `_carregar_historico()`, que continua privado deste módulo (quem
    lê o arquivo em disco é só aqui).
    """
    return _carregar_historico()


def meses_disponiveis() -> List[str]:
    """
    Lista, do mais recente para o mais antigo, os meses "MM/AAAA" com ao
    menos 1 execução registrada - usada para popular o seletor de mês da
    aba Relatório. O mês atual sempre aparece, mesmo sem nenhuma
    automação executada ainda nele.
    """
    historico = _carregar_historico()
    meses = set()
    for item in historico:
        try:
            meses.add(datetime.fromisoformat(item["inicio"]).strftime("%m/%Y"))
        except (KeyError, ValueError):
            continue
    meses.add(datetime.now().strftime("%m/%Y"))
    return sorted(meses, key=lambda m: (m[3:], m[:2]), reverse=True)


def _bloco_vazio() -> dict:
    return {"execucoes": 0, "processados": 0, "duracao_total_segundos": 0, "duracao_media_segundos": 0}


def _acumular(bloco: dict, item: dict) -> None:
    bloco["execucoes"] += 1
    bloco["processados"] += item.get("processados", 0)
    bloco["duracao_total_segundos"] += item.get("duracao_segundos", 0)


def _mesmo_mes(inicio_iso: str, mes: int, ano: int) -> bool:
    try:
        dt = datetime.fromisoformat(inicio_iso)
    except (TypeError, ValueError):
        return False
    return dt.month == mes and dt.year == ano


def resumo_mes(mes_ano: str) -> dict:
    """
    Filtra o histórico para o mês/ano informado ("MM/AAAA" - ex:
    "09/2026"), do 1º ao último dia do mês (inclusive - basta comparar
    mês e ano de cada registro, sem precisar calcular o último dia à
    mão), e devolve um resumo pronto para a aba Relatório:

        {
            "mes_ano": "09/2026",
            "categorias": {"CE": {...}, "NL": {...}, "PP": {...}, "OB": {...}},
            "detalhamento": [ {"programa": "CE", ...}, ... - 1 por tipo de automação ],
            "total_execucoes": int,
        }

    Cada bloco (de categoria ou de automação) traz: execucoes,
    processados, duracao_total_segundos, duracao_media_segundos (tempo
    total dividido pelos itens processados - 0 se nada foi processado).
    """
    try:
        mes_texto, ano_texto = mes_ano.split("/")
        mes, ano = int(mes_texto), int(ano_texto)
    except (ValueError, AttributeError):
        agora = datetime.now()
        mes, ano = agora.month, agora.year

    historico = _carregar_historico()
    do_mes = [item for item in historico if _mesmo_mes(item.get("inicio"), mes, ano)]

    categorias = {categoria: _bloco_vazio() for categoria in CATEGORIAS_PRINCIPAIS}
    por_programa = {}

    for item in do_mes:
        categoria = item.get("categoria")
        if categoria in categorias:
            _acumular(categorias[categoria], item)

        programa = item.get("programa", "?")
        por_programa.setdefault(programa, _bloco_vazio())
        _acumular(por_programa[programa], item)

    for bloco in list(categorias.values()) + list(por_programa.values()):
        bloco["duracao_media_segundos"] = (
            bloco["duracao_total_segundos"] / bloco["processados"] if bloco["processados"] else 0
        )

    return {
        "mes_ano": f"{mes:02d}/{ano}",
        "categorias": categorias,
        "detalhamento": [
            {"programa": programa, **bloco} for programa, bloco in sorted(por_programa.items())
        ],
        "total_execucoes": len(do_mes),
    }


def formatar_duracao(segundos: float) -> str:
    """Converte segundos num texto tipo '1h 12min 03s' ou '45s' - mais
    fácil de ler do que um número de segundos cru para quem não é
    técnico. `0` (ou vazio) vira '0s', nunca texto em branco."""
    segundos = int(round(segundos or 0))
    horas, resto = divmod(segundos, 3600)
    minutos, seg = divmod(resto, 60)
    partes = []
    if horas:
        partes.append(f"{horas}h")
    if horas or minutos:
        partes.append(f"{minutos}min")
    partes.append(f"{seg}s")
    return " ".join(partes)
