"""Automação Gerar OB (Ordem Bancária, lote de 30 em 30), Raspar OB e Confirmar OB."""
from typing import List, Optional

from .. import colunas
from ..constantes import (
    URL_OB_SIGEF, MARCADOR_URL_OB, TIMEOUT_PADRAO_SIGEF, TIMEOUT_CLIQUE_GRADE,
    TENTATIVAS_CLIQUE_GRADE, TAMANHO_LOTE_OB, MAX_PAGINAS_LOTE_OB, REGEX_APENAS_DIGITOS,
    REGEX_DOCUMENTO_OB,
)
from ..excel import obter_celula, atualizar_status, salvar_valor_gerado
from ..log import log_info, log_sucesso, log_erro, log_aviso
from ..navegador import conectar_e_obter_pagina_sigef, aguardar_pagina_estavel, fechar_paginas, abrir_popup
from ..playwright_compat import sync_playwright, Page, BrowserContext
from ..utils import formatar_cpf, formatar_valor_br
from .nl import _valor_br_para_float
from .pp import _valor_planilha_para_float

def _cpf_para_comparacao_ob(valor) -> str:
    """
    Normaliza um CPF (vindo da planilha OU da grade do SIGEF) para 11
    dígitos, sem pontuação/máscara, com zeros à esquerda preservados -
    usado só para COMPARAR os dois lados. Para digitar num campo do
    SIGEF, use `formatar_cpf` (devolve com máscara XXX.XXX.XXX-XX).
    """
    digitos = REGEX_APENAS_DIGITOS.sub("", str(valor or ""))
    return digitos.zfill(11) if digitos else ""


def _ler_grade_ob(popup: "Page") -> List[dict]:
    """
    Lê a grade `#dtgPrepPgto` (Preparação de Pagamento) inteira em UMA
    chamada JS, em vez de 1 chamada Playwright por célula. Índices de
    célula (6=PP, 13=Favorecido/CPF, 15=Valor) conferidos no HTML da
    tela "Manter Ordem Bancária" > "Adicionar" (mesmo esquema usado pelo
    `_localizar_registro_pp` do pacote `sigef_automacao`, que só
    precisava de PP e Valor).
    """
    return popup.evaluate(
        """
        () => {
            const linhas = document.querySelectorAll("tr.GridLinhaPar, tr.GridLinhaImpar");
            return [...linhas].map((linha, indice) => {
                const celulas = [...linha.querySelectorAll("td")].map(td => td.innerText.trim());
                return { indice: indice, pp: celulas[6] || "", cpf: celulas[13] || "", valor: celulas[15] || "" };
            });
        }
        """
    )


def _buscar_item_na_grade_ob(grade: List[dict], item: dict) -> Optional[dict]:
    """
    Só considera um registro da grade como "o item da planilha" quando os
    TRÊS critérios batem ao mesmo tempo: PP, CPF do favorecido e Valor -
    todos padronizados antes de comparar, reaproveitando as MESMAS
    funções já usadas por Raspar PP/Raspar contas para os mesmos dados:
    `_valor_br_para_float` (valor da grade, com separador BR) e
    `_valor_planilha_para_float` (valor da planilha, coluna H, em reais -
    NÃO usar `normalizar_valor_sigef` aqui, que trataria um "1330" da
    planilha como centavos e devolveria 13,30 em vez de 1.330,00).
    """
    valor_esperado = _valor_planilha_para_float(item["valor"])
    cpf_esperado = _cpf_para_comparacao_ob(item["cpf_esperado"])

    for linha in grade:
        if linha["pp"] != item["pp_esperado"]:
            continue
        if _cpf_para_comparacao_ob(linha["cpf"]) != cpf_esperado:
            continue
        try:
            valor_grade = _valor_br_para_float(linha["valor"])
        except (ValueError, TypeError):
            continue
        if abs(valor_grade - valor_esperado) < 0.01:
            return linha
    return None


def _marcar_checkbox_ob(popup: "Page", indice: int) -> bool:
    """Marca `#chk{indice}` da grade, com retentativas - a grade às vezes
    engole o 1º clique (mesmo comportamento já visto em outras telas do
    SIGEF, ver `TENTATIVAS_CLIQUE_GRADE`)."""
    checkbox = popup.locator(f"#chk{indice}")
    for _ in range(TENTATIVAS_CLIQUE_GRADE):
        try:
            checkbox.wait_for(state="visible", timeout=TIMEOUT_CLIQUE_GRADE)
            if checkbox.is_checked():
                return True
            checkbox.click()
            popup.wait_for_timeout(150)
            if checkbox.is_checked():
                return True
        except Exception:
            pass
    log_aviso(f"Não foi possível confirmar a marcação de #chk{indice} após {TENTATIVAS_CLIQUE_GRADE} tentativa(s).")
    return False


def _avancar_pagina_ob(popup: "Page") -> bool:
    """Clica em 'Próxima página' (`#pagFormulario_BtnProximo`). Retorna
    False se o botão não existir ou o clique falhar - tratado como 'não
    há mais páginas' por `_montar_lote_ob`."""
    botao = popup.locator("#pagFormulario_BtnProximo")
    if botao.count() == 0:
        return False
    try:
        botao.click()
    except Exception:
        return False

    aguardar_pagina_estavel(popup)
    try:
        popup.wait_for_selector("tr.GridLinhaPar, tr.GridLinhaImpar", timeout=TIMEOUT_PADRAO_SIGEF)
    except Exception:
        return False
    return True


def _anotar_falha_ob(worksheet, item: dict, motivo: str) -> None:
    """Grava o motivo da falha na coluna L da linha (mesma coluna usada
    para o número da OB quando dá certo) - a planilha fica com o
    histórico completo de cada linha, tenha ela entrado numa OB ou não."""
    mensagem = (
        f"{motivo} (PP {item['pp_esperado']}, CPF {formatar_cpf(item['cpf_esperado'])}, "
        f"valor {formatar_valor_br(item['valor'])})"
    )
    log_erro(f"Linha {item['numero_linha']}: {mensagem}")
    atualizar_status(worksheet, item["numero_linha"], colunas.COL_OB_GERADA, mensagem)


def _abastecer_lote_ob(dados: List, cursor: int, linha_inicial: int,
                        pendentes_atual: List[dict], faltam: int) -> tuple:
    """
    Garante que `pendentes_atual` tenha pelo menos `faltam` itens,
    puxando novas linhas VÁLIDAS (com PP, CPF e valor preenchidos) da
    planilha a partir de `cursor`. Linhas inválidas (algum campo vazio)
    são só puladas - não contam para o lote nem são anotadas como falha.
    """
    pendentes = list(pendentes_atual)
    while len(pendentes) < faltam and cursor < len(dados):
        linha = dados[cursor]
        numero_linha = linha_inicial + cursor
        cursor += 1

        pp_esperado = obter_celula(linha, colunas.COL_OB_PP_ESPERADO)
        cpf_esperado = obter_celula(linha, colunas.COL_OB_CPF)
        valor_bruto = obter_celula(linha, colunas.COL_OB_VALOR)
        if not pp_esperado or not cpf_esperado or not valor_bruto:
            log_aviso(f"Linha {numero_linha}: PP (coluna K), CPF (coluna B) ou valor "
                      f"(coluna H) vazio. Pulando.")
            continue

        pendentes.append({
            "numero_linha": numero_linha,
            "pp_esperado": pp_esperado,
            "cpf_esperado": cpf_esperado,
            "valor": valor_bruto,
        })
    return pendentes, cursor


def _montar_lote_ob(popup: "Page", dados: List, cursor: int, linha_inicial: int, worksheet) -> tuple:
    """
    Monta e marca (via checkbox) até `TAMANHO_LOTE_OB` itens confirmados
    na grade `popup`, avançando páginas (`_avancar_pagina_ob`) e repondo
    com novas linhas da planilha (`_abastecer_lote_ob`) sempre que um
    item não é encontrado - assim um lote só fecha com MENOS itens que
    `TAMANHO_LOTE_OB` se a planilha realmente acabar ou a busca esgotar
    as páginas disponíveis.

    Retorna (confirmados, cursor_atualizado). `confirmados` é a lista de
    dicts (mesma forma de `_abastecer_lote_ob`) cujo checkbox já foi
    marcado na grade.

    LOOP: laço de BUSCA/PAGINAÇÃO deste lote (`while True`), com 5
    condições de saída - nunca infinito, porque a cada volta ou marca
    item(ns), ou avança de página, ou encerra por uma das condições
    abaixo:
        1. já confirmou os `TAMANHO_LOTE_OB` itens do lote;
        2. a planilha acabou e não sobrou item pendente para repor;
        3. não há botão de próxima página (ou o clique falhou);
        4. a página seguinte é IDÊNTICA a uma já vista (a paginação deu
           volta ou parou de avançar de verdade);
        5. limite de segurança `MAX_PAGINAS_LOTE_OB` atingido.
    """
    confirmados: List[dict] = []
    pendentes, cursor = _abastecer_lote_ob(dados, cursor, linha_inicial, [], TAMANHO_LOTE_OB)
    paginas_vistas = set()

    while True:
        try:
            popup.wait_for_selector("tr.GridLinhaPar, tr.GridLinhaImpar", timeout=TIMEOUT_PADRAO_SIGEF)
        except Exception:
            log_erro("A grade de itens (Preparação de Pagamento) não apareceu na tela.")
            for item in pendentes:
                _anotar_falha_ob(worksheet, item, "Grade de itens não carregou.")
            break

        grade = _ler_grade_ob(popup)

        assinatura = tuple((linha["pp"], linha["valor"]) for linha in grade)
        if assinatura in paginas_vistas:
            log_aviso("A grade repetiu uma página já vista (fim da paginação). Encerrando busca deste lote.")
            for item in pendentes:
                _anotar_falha_ob(worksheet, item, "PP não localizado em nenhuma página da grade.")
            break
        paginas_vistas.add(assinatura)

        ainda_pendentes = []
        for item in pendentes:
            registro = _buscar_item_na_grade_ob(grade, item)
            if registro is None:
                ainda_pendentes.append(item)
                continue
            _marcar_checkbox_ob(popup, registro["indice"])
            confirmados.append(item)
        pendentes = ainda_pendentes

        if len(confirmados) >= TAMANHO_LOTE_OB:
            cursor -= len(pendentes)  # devolve o que sobrou (não tentado) para o próximo lote
            break

        pendentes, cursor = _abastecer_lote_ob(
            dados, cursor, linha_inicial, pendentes, TAMANHO_LOTE_OB - len(confirmados)
        )

        if not pendentes:
            break  # nada mais a procurar (planilha acabou e não sobrou pendente)

        if not _avancar_pagina_ob(popup):
            for item in pendentes:
                _anotar_falha_ob(worksheet, item, "PP não localizado em nenhuma página da grade.")
            break

        if len(paginas_vistas) >= MAX_PAGINAS_LOTE_OB:
            log_aviso(f"Limite de {MAX_PAGINAS_LOTE_OB} páginas atingido neste lote.")
            for item in pendentes:
                _anotar_falha_ob(worksheet, item, "Limite de páginas de busca atingido.")
            break

    return confirmados, cursor


def _preencher_cabecalho_ob(page: "Page", config: dict) -> None:
    """
    Preenche o cabeçalho da tela "Manter Ordem Bancária" - fluxo
    Descentralizada (Tipo de OB "2"), os mesmos valores do OB_amigo.py
    original: Gestão, Banco/Agência/Conta de origem, Tipo de OB, Tipo de
    Pagamento e a Observação com o mês de referência (agora vindo de
    verdade de `config["mes_referencia"]`, o mesmo campo já usado pelas
    demais automações e configurado no menu "Configurar parâmetros").
    """
    mes_referencia = config.get("mes_referencia", "")
    processo = config.get("processo", "")

    page.locator("#txtGestao_SIGEFPesquisa").fill("0001")
    page.locator("#txtBancoOrigem").fill("001")
    page.locator("#txtAgenciaOrigem").fill("2757X")
    page.locator("#txtContaOrigem").fill(config.get("conta_origem") or "100005")
    page.locator("#cboTipoOB").select_option(value="2")        # Descentralizada
    page.locator("#cboTipoPagamento").select_option(value="2")
    page.locator("#txtDeObservacao").fill(
        f"Ressarcimento de amigos voluntario Referente ao {mes_referencia} Processo: {processo}"
    )


def _garantir_cabecalho_ob(page: "Page", config: dict) -> None:
    """
    Preenche o cabeçalho da OB apenas 1 ÚNICA VEZ, quando a tela é
    aberta: os campos permanecem preenchidos entre um lote e outro porque
    `_limpar_para_proximo_lote_ob` só mexe na grade de itens incluídos
    (Selecionar Todos + Remover), nunca no cabeçalho.

    Chamada antes de CADA lote, mas só preenche de fato quando a Gestão
    está vazia - na 1ª chamada (sempre) e, se algum imprevisto do SIGEF
    limpar a tela no meio da execução (ex: sessão expirou e a página
    recarregou sozinha), nas chamadas seguintes também: essa é a "regra
    de continuação" pedida por você - o cabeçalho só é refeito se, na
    prática, precisar ser refeito.
    """
    campo_gestao = page.locator("#txtGestao_SIGEFPesquisa")
    try:
        if campo_gestao.count() and campo_gestao.input_value().strip():
            return
    except Exception:
        pass
    log_info("Preenchendo o cabeçalho da OB (Gestão, Banco, Agência, Conta, Tipo de OB/Pagamento)...")
    _preencher_cabecalho_ob(page, config)


def _abrir_popup_pesquisa_ob(page: "Page", context: "BrowserContext") -> "Page":
    """Clica em 'Adicionar' e preenche a pesquisa do popup (Unidade
    Gestora, Gestão, Id de Uso e Fonte) - mesmos 4 campos para TODOS os
    lotes, refeitos a cada popup novo (o SIGEF não guarda isso de um
    popup para o outro)."""
    popup = abrir_popup(context, lambda: page.locator("#btnAdicionar").click())

    popup.locator("#txtUnidadeGestora").fill("160001")
    popup.locator("#txtGestao_SIGEFPesquisa").fill("0001")
    popup.locator("#txtIdUso").fill("1")
    popup.locator("#txtFonte_SIGEFPesquisa").fill("500001001")
    popup.locator("#btnPesquisar").click()
    aguardar_pagina_estavel(popup)
    return popup


def _confirmar_lote_no_popup_ob(popup: "Page") -> None:
    botao_confirmar = popup.locator("#btnConfirmar")
    botao_confirmar.wait_for(state="visible", timeout=TIMEOUT_PADRAO_SIGEF)
    botao_confirmar.click()
    popup.wait_for_event("close", timeout=TIMEOUT_PADRAO_SIGEF)


def _confirmar_ob_e_capturar_numero(page: "Page") -> Optional[str]:
    """
    Clica em "Incluir" na tela principal e devolve o número da OB gerada
    - ou None se o SIGEF recusou o lote (mensagem de erro) ou o número
    não pôde ser extraído; o motivo, em qualquer um dos dois casos, já
    vai para o log (mesmo padrão da automação PP)."""
    page.bring_to_front()
    aguardar_pagina_estavel(page)

    botao_incluir = page.locator("#SIGEFBotoesManutencao_BtnIncluir")
    botao_incluir.wait_for(state="visible", timeout=TIMEOUT_PADRAO_SIGEF)
    botao_incluir.click()

    page.wait_for_selector(
        "td.SIGEFMensagemSucesso, td.SIGEFMensagemErro", timeout=TIMEOUT_PADRAO_SIGEF
    )

    if page.locator("td.SIGEFMensagemErro").count() > 0:
        mensagem_erro = page.locator("td.SIGEFMensagemErro").first.inner_text().strip()
        log_erro(f"SIGEF recusou a OB deste lote: {mensagem_erro}")
        return None

    mensagem_sucesso = page.locator("td.SIGEFMensagemSucesso").first.inner_text().strip()
    busca = REGEX_DOCUMENTO_OB.search(mensagem_sucesso)
    if busca is None:
        log_erro(f"Não foi possível extrair o número da OB da mensagem do SIGEF ({mensagem_sucesso}).")
        return None
    return busca.group(0).upper()


def _limpar_para_proximo_lote_ob(page: "Page") -> None:
    """
    Remove os itens incluídos da lista da tela principal, preparando para
    o próximo lote. Usa 'Selecionar Todos' (`#Selecao_SelectAll`) em vez
    de marcar só o 1º checkbox (como no fluxo de 1 item por OB do pacote
    `sigef_automacao`/`ordem_bancaria.py`), porque aqui um lote inclui
    vários itens de uma vez. NUNCA mexe no cabeçalho (Gestão, Banco,
    Tipo de OB...) - só na grade de itens incluídos.
    """
    try:
        page.locator("#txtNumeroOB").fill("")
        page.locator("#Selecao_SelectAll").check()
        page.locator("#btnRemover").click()
    except Exception as erro:
        log_aviso(f"Não foi possível limpar a tela para o próximo lote: {erro}")


def gerar_ob(dados, config=None, worksheet=None):
    """
    Automação: Gerar OB (lote de 30 em 30) - tela "Manter Ordem Bancária",
    fluxo Descentralizada (Tipo de OB "2").

    Em vez de abrir 1 popup de pesquisa por linha da planilha (fluxo de
    Regularização do pacote `sigef_automacao`/`ordem_bancaria.py`, 1 PP =
    1 OB), esta automação:
        1. Conecta 1 única vez à tela de Ordem Bancária e garante o
           cabeçalho preenchido (Gestão, Banco, Agência, Conta, Tipo de
           OB, Tipo de Pagamento, Observação com o mês de referência) -
           ver `_garantir_cabecalho_ob`.
        2. Para cada lote de até `TAMANHO_LOTE_OB` linhas da planilha:
           abre 1 popup de pesquisa e procura na grade `#dtgPrepPgto`
           (paginando com `_avancar_pagina_ob` quando necessário) os
           itens cujo PP (coluna K), CPF do favorecido (coluna B) e
           Valor (coluna H) batem com a planilha, marcando o checkbox de
           cada um (`_montar_lote_ob`).
        3. Item não encontrado em NENHUMA página é anotado na coluna L
           com o motivo, e o lote é REPOSTO com a próxima linha ainda
           não tentada da planilha, para fechar em `TAMANHO_LOTE_OB`
           itens confirmados mesmo assim (ver `_abastecer_lote_ob`).
        4. Confirma o lote (Confirmar no popup + Incluir na tela
           principal), captura o número da OB gerada e grava na coluna L
           de TODAS as linhas confirmadas daquele lote.
        5. Limpa a lista de itens incluídos (sem mexer no cabeçalho) e
           repete para o próximo lote, até acabar a planilha.

    LIMITAÇÃO CONHECIDA: a grade só tem botão de avançar página
    (`#pagFormulario_BtnProximo`) - sem "anterior"/"primeira página" no
    HTML fornecido. Por isso a reposição só encontra o item de troca se
    o PP dele estiver numa página AINDA NÃO visitada nesta rodada; se
    estiver numa página já passada, também é anotado como não
    encontrado (no pior caso, um lote fecha com menos itens do que
    `TAMANHO_LOTE_OB`).

    LOOP PRINCIPAL: o `while cursor < len(dados)` mais abaixo percorre a
    planilha INTEIRA, lote de 30 em 30, gerando 1 OB por lote, até não
    sobrar mais linha - nunca infinito, porque `_montar_lote_ob` sempre
    avança `cursor` (confirmando itens e/ou consumindo linhas via
    reposição), mesmo no pior caso em que nada é encontrado em lugar
    nenhum.
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return

    if not dados:
        log_aviso("Nenhuma linha para processar na automação Gerar OB.")
        return

    config = config or {}
    linha_inicial = config.get("linha_inicial", 2)

    with sync_playwright() as p:
        try:
            context, page = conectar_e_obter_pagina_sigef(p, URL_OB_SIGEF, MARCADOR_URL_OB)
        except Exception as erro:
            log_erro(f"Erro ao conectar à tela de Ordem Bancária do SIGEF: {erro}")
            return

        cursor = 0
        total_obs = 0
        total_itens_confirmados = 0

        while cursor < len(dados):
            # "Regra de continuação": só preenche de novo se necessário
            # (ver docstring de `_garantir_cabecalho_ob`).
            _garantir_cabecalho_ob(page, config)

            try:
                popup = _abrir_popup_pesquisa_ob(page, context)
            except Exception as erro:
                log_erro(f"Não foi possível abrir a pesquisa de lançamentos: {erro}")
                break

            confirmados, cursor = _montar_lote_ob(popup, dados, cursor, linha_inicial, worksheet)

            if not confirmados:
                log_aviso("Nenhum item confirmado neste lote; nada para submeter.")
                fechar_paginas([popup])
                page.bring_to_front()
                continue

            try:
                _confirmar_lote_no_popup_ob(popup)
            except Exception as erro:
                log_erro(f"Falha ao confirmar o lote no popup de pesquisa: {erro}")
                for item in confirmados:
                    _anotar_falha_ob(worksheet, item, f"Falha ao confirmar o lote no popup: {erro}")
                page.bring_to_front()
                continue

            numero_ob = _confirmar_ob_e_capturar_numero(page)
            if numero_ob is None:
                for item in confirmados:
                    _anotar_falha_ob(worksheet, item, "SIGEF recusou ou não confirmou a OB deste lote.")
                continue

            for item in confirmados:
                salvar_valor_gerado(worksheet, item["numero_linha"], colunas.COL_OB_GERADA, numero_ob, rotulo="OB")

            total_obs += 1
            total_itens_confirmados += len(confirmados)
            log_sucesso(f"OB '{numero_ob}' gerada com {len(confirmados)} item(ns) do lote.")

            _limpar_para_proximo_lote_ob(page)

        log_sucesso(
            f"Automação 'Gerar OB' concluída: {total_obs} OB(s) gerada(s), "
            f"{total_itens_confirmados} item(ns) confirmados de {len(dados)} linha(s) da planilha."
        )


def raspar_ob(dados, config=None, worksheet=None):
    """Automação: Raspar OB (a implementar)."""
    log_info("Executando: Raspar OB (não implementado ainda).")


def confirmar_ob(dados, config=None, worksheet=None):
    """Automação: Confirmar OB (a implementar)."""
    log_info("Executando: Confirmar OB (não implementado ainda).")


