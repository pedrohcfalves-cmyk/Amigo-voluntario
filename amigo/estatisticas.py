# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
ESTATÍSTICAS
============
Análises sobre o histórico de execuções (amigo/relatorio.py ->
historico.json): quanto tempo a automação economiza frente ao trabalho
manual equivalente, e como esse tempo se comporta em função do número de
itens processados. Alimenta a aba "Estatísticas" da interface gráfica,
pensada para reunir dados que embasem um estudo de implementação de
automações em larga escala numa secretaria (não só o uso individual do
Amigo) - por isso o foco não é só "quanto já rodou", mas também "o que dá
para prever se o volume crescer".

Este módulo NÃO inventa um número "absoluto" de tempo manual: usa uma
estimativa de minutos por item, por categoria (CE/NL/PP/OB/Conferência),
que é um PARÂMETRO configurável (salvo em config.json, chave
"tempo_manual_estimado_minutos") e ajustável na própria aba "Estatísticas"
- assim o cálculo reflete a realidade observada no setor, em vez de um
número fixo genérico. Os valores de fábrica (abaixo) são só um ponto de
partida - tempo típico de preencher/confirmar 1 tela do SIGEF à mão
(navegar, digitar, revisar, submeter) - e devem ser ajustados por quem
conhece a rotina manual real.
"""
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from .config_io import salvar_configuracoes
from .relatorio import CATEGORIAS_PRINCIPAIS, historico_completo

# Categorias com que a aba Estatísticas trabalha: as 4 principais (que
# geram documento) + Conferência (Raspar contas) - não gera documento
# novo, mas também substitui um trabalho manual (confere dados na tela) e
# por isso também entra na conta de tempo economizado.
CATEGORIAS_ESTATISTICAS = CATEGORIAS_PRINCIPAIS + ["Conferência"]

# Estimativa de fábrica: minutos gastos para fazer 1 item À MÃO (sem o
# Amigo) - navegar até a tela do SIGEF, digitar os dados, revisar e
# confirmar. É AJUSTÁVEL na aba "Estatísticas" - o valor real varia de
# pessoa para pessoa e de setor para setor.
TEMPO_MANUAL_PADRAO_MINUTOS = {
    "CE": 3.0,
    "NL": 3.0,
    "PP": 3.0,
    "OB": 4.0,
    "Conferência": 2.0,
}

CHAVE_CONFIG_TEMPO_MANUAL = "tempo_manual_estimado_minutos"

# Volumes de itens usados na projeção "e se o uso crescesse?" - pensados
# para uma secretaria inteira (não só 1 usuário), do tamanho de um lote
# hoje até uma operação de larga escala.
VOLUMES_PROJECAO = [100, 500, 1_000, 5_000, 10_000, 20_000]


def obter_tempos_manuais(config: dict) -> Dict[str, float]:
    """
    Devolve o dict {categoria: minutos por item} usado para estimar o
    tempo manual - mescla o que está salvo em config.json com o padrão de
    fábrica, para que uma categoria nova (numa atualização futura) sempre
    tenha valor, e para que o usuário possa ter ajustado só ALGUMAS
    categorias, mantendo as outras no padrão.
    """
    tempos = dict(TEMPO_MANUAL_PADRAO_MINUTOS)
    salvos = config.get(CHAVE_CONFIG_TEMPO_MANUAL)
    if isinstance(salvos, dict):
        for categoria, minutos in salvos.items():
            try:
                tempos[categoria] = max(0.0, float(minutos))
            except (TypeError, ValueError):
                continue
    return tempos


def salvar_tempos_manuais(config: dict, tempos: Dict[str, float]) -> None:
    """Grava as estimativas de tempo manual (minutos/item) em config.json,
    junto do restante da configuração - chamada pelo botão 'Salvar
    estimativas' da aba Estatísticas."""
    config[CHAVE_CONFIG_TEMPO_MANUAL] = {
        categoria: float(minutos) for categoria, minutos in tempos.items()
    }
    salvar_configuracoes(config)


def _historico_no_periodo(periodo: Optional[str]) -> List[dict]:
    """`periodo` é 'MM/AAAA' ou `None`/'Todo o período' para não filtrar."""
    historico = historico_completo()
    if not periodo or periodo == "Todo o período":
        return historico
    try:
        mes_texto, ano_texto = periodo.split("/")
        mes, ano = int(mes_texto), int(ano_texto)
    except (ValueError, AttributeError):
        return historico

    filtrado = []
    for item in historico:
        try:
            data = datetime.fromisoformat(item["inicio"])
        except (KeyError, ValueError, TypeError):
            continue
        if data.month == mes and data.year == ano:
            filtrado.append(item)
    return filtrado


def _regressao_linear(pontos: List[Tuple[int, float]]) -> Optional[Tuple[float, float]]:
    """
    Regressão linear simples (mínimos quadrados) da duração (segundos) em
    função dos itens processados, em 1 execução: duração ≈ intercepto +
    inclinação × itens.

    O INTERCEPTO é o "custo fixo" de rodar a automação 1 vez (login,
    navegação, confirmação) - tempo que não depende de quantos itens
    entraram naquela execução -, e a INCLINAÇÃO é o custo marginal por
    item (quanto cada item adicional soma ao tempo total). Essa separação
    é o dado mais útil para pensar escala: mostra se compensa mais rodar
    poucas automações grandes (custo fixo se dilui) do que muitas
    pequenas.

    Devolve `None` com menos de 2 execuções com itens diferentes (não dá
    para ajustar uma reta) - quem chama cai de volta para a média simples.
    """
    if len(pontos) < 2 or len({x for x, _y in pontos}) < 2:
        return None
    n = len(pontos)
    soma_x = sum(x for x, _y in pontos)
    soma_y = sum(y for _x, y in pontos)
    soma_xy = sum(x * y for x, y in pontos)
    soma_xx = sum(x * x for x, _y in pontos)
    denominador = n * soma_xx - soma_x * soma_x
    if denominador == 0:
        return None
    inclinacao = (n * soma_xy - soma_x * soma_y) / denominador
    intercepto = (soma_y - inclinacao * soma_x) / n
    return intercepto, inclinacao


def _bloco_vazio() -> dict:
    return {
        "execucoes": 0,
        "processados": 0,
        "duracao_automacao_segundos": 0.0,
        "_pontos_regressao": [],
    }


def calcular_estatisticas(periodo: Optional[str], config: dict) -> dict:
    """
    Monta todos os números da aba Estatísticas para o período escolhido
    ('MM/AAAA' ou 'Todo o período'):

        {
            "periodo": "...",
            "tempos_manuais_minutos": {categoria: minutos_por_item},
            "por_categoria": {categoria: {
                "execucoes", "processados",
                "duracao_automacao_segundos", "duracao_manual_segundos",
                "tempo_economizado_segundos", "percentual_economia",
                "segundos_por_item_automacao",
                "regressao": (intercepto_segundos, inclinacao_segundos_por_item) | None,
            }},
            "totais": {mesmas chaves, somadas de todas as categorias},
            "projecao": [
                {"itens", "tempo_automatizado_segundos", "tempo_manual_segundos",
                 "tempo_economizado_segundos"},
                ...  # 1 por volume em VOLUMES_PROJECAO
            ],
        }

    A projeção usa a taxa média GERAL observada no período (segundos por
    item, somando todas as categorias) - é uma simplificação proposital:
    numa secretaria real, o "próximo lote" de itens tende a ser uma mistura
    de CE/NL/PP/OB, não só de 1 tipo.
    """
    historico = _historico_no_periodo(periodo)
    tempos_manuais = obter_tempos_manuais(config)

    por_categoria = {categoria: _bloco_vazio() for categoria in CATEGORIAS_ESTATISTICAS}

    for item in historico:
        categoria = item.get("categoria")
        bloco = por_categoria.get(categoria)
        if bloco is None:
            continue
        bloco["execucoes"] += 1
        processados = int(item.get("processados") or 0)
        bloco["processados"] += processados
        bloco["duracao_automacao_segundos"] += item.get("duracao_segundos", 0)
        if processados:
            bloco["_pontos_regressao"].append((processados, item.get("duracao_segundos", 0)))

    for categoria, bloco in por_categoria.items():
        minutos_por_item = tempos_manuais.get(categoria, 0.0)
        bloco["duracao_manual_segundos"] = bloco["processados"] * minutos_por_item * 60
        bloco["tempo_economizado_segundos"] = (
            bloco["duracao_manual_segundos"] - bloco["duracao_automacao_segundos"]
        )
        bloco["percentual_economia"] = (
            (bloco["tempo_economizado_segundos"] / bloco["duracao_manual_segundos"]) * 100
            if bloco["duracao_manual_segundos"] else 0.0
        )
        bloco["segundos_por_item_automacao"] = (
            bloco["duracao_automacao_segundos"] / bloco["processados"] if bloco["processados"] else 0.0
        )
        bloco["regressao"] = _regressao_linear(bloco.pop("_pontos_regressao"))

    totais = {
        "execucoes": sum(b["execucoes"] for b in por_categoria.values()),
        "processados": sum(b["processados"] for b in por_categoria.values()),
        "duracao_automacao_segundos": sum(b["duracao_automacao_segundos"] for b in por_categoria.values()),
        "duracao_manual_segundos": sum(b["duracao_manual_segundos"] for b in por_categoria.values()),
    }
    totais["tempo_economizado_segundos"] = (
        totais["duracao_manual_segundos"] - totais["duracao_automacao_segundos"]
    )
    totais["percentual_economia"] = (
        (totais["tempo_economizado_segundos"] / totais["duracao_manual_segundos"]) * 100
        if totais["duracao_manual_segundos"] else 0.0
    )
    totais["segundos_por_item_automacao"] = (
        totais["duracao_automacao_segundos"] / totais["processados"] if totais["processados"] else 0.0
    )
    media_manual_segundos_por_item = (
        totais["duracao_manual_segundos"] / totais["processados"] if totais["processados"] else 0.0
    )

    taxa_automacao = totais["segundos_por_item_automacao"]
    projecao = []
    for itens in VOLUMES_PROJECAO:
        tempo_automatizado = itens * taxa_automacao
        tempo_manual = itens * media_manual_segundos_por_item
        projecao.append({
            "itens": itens,
            "tempo_automatizado_segundos": tempo_automatizado,
            "tempo_manual_segundos": tempo_manual,
            "tempo_economizado_segundos": tempo_manual - tempo_automatizado,
        })

    return {
        "periodo": periodo or "Todo o período",
        "tempos_manuais_minutos": tempos_manuais,
        "por_categoria": por_categoria,
        "totais": totais,
        "projecao": projecao,
    }
