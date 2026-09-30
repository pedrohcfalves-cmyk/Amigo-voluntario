# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""Automação PP (Preparação de Pagamento) e Raspar PP."""
from datetime import datetime

from .. import colunas
from ..constantes import (
    URL_PP_SIGEF, MARCADOR_URL_PP, URL_RASPAR_CONTA_SIGEF, MARCADOR_URL_RASPAR_CONTA,
    TIMEOUT_PADRAO_SIGEF, TIMEOUT_POPUP_FECHAR, TIMEOUT_CLIQUE_GRADE, TENTATIVAS_CLIQUE_GRADE,
    TIPO_ORDEM_BANCARIA_PP, REGEX_APENAS_DIGITOS, REGEX_PREFIXO_DOCUMENTO, REGEX_DOCUMENTO_PP,
    REGEX_NUMEROS, REGEX_VALOR_BR, ano_do_exercicio, url_do_exercicio,
)
from ..execucao import (
    eh_simulado, em_simulacao, informar_progresso, linha_em_execucao, parar_antes_da_linha,
)
from ..excel import obter_celula, atualizar_status, salvar_valor_gerado
from ..log import log_info, log_sucesso, log_erro, log_aviso
from ..navegador import (
    conectar_e_obter_pagina_sigef, aguardar_pagina_estavel, aguardar_pagina_ociosa,
    clicar_e_aguardar_elemento, limpar_formulario, aguardar_campo_com_valor_por_id,
    selecionar_combo_com_conferencia, fechar_paginas, abrir_popup,
)
from ..playwright_compat import sync_playwright, Page, BrowserContext, PlaywrightTimeoutError, PlaywrightError
from ..utils import (
    extrair_numero_documento, normalizar_numero_documento,
    normalizar_chave_bancaria, formatar_cpf, formatar_valor_centavos, formatar_valor_br,
)
from .ce import _aguardar_resultado_ou_erro_grade
from .nl import _valor_br_para_float

def ler_tabela_domicilio_bancario(popup: "Page"):
    """
    Lê TODA a grade `#dtgDomicilioBancario` de uma única vez, via
    JavaScript (`evaluate`) - mesmo padrão de
    `ler_tabela_despesa_certificada`: 1 round-trip em vez de 1 chamada ao
    Playwright por célula (a leitura célula a célula era o trecho mais
    lento da PP).

    Devolve apenas as linhas selecionáveis (`td.GridLink`), já sem o
    cabeçalho e sem as linhas vazias (`&nbsp;`) do rodapé da grade. O
    campo `indice` é a posição da linha entre os `td.GridLink` da grade,
    usada depois para clicar sem varrer o DOM de novo.
    """
    return popup.locator("#dtgDomicilioBancario").evaluate("""
        table => [...table.querySelectorAll("td.GridLink")].map((tdBanco, indice) => {
            const celulas = [...tdBanco.closest("tr").querySelectorAll("td")]
                .slice(0, 3)
                .map(td => td.textContent.trim());

            return {
                indice: indice,
                banco: celulas[0] || "",
                agencia: celulas[1] || "",
                conta: celulas[2] || ""
            };
        })
    """)


def ler_tabela_ordem_cronologica(popup: "Page"):
    """
    Lê a grade "Gerar Ordem Cronológica" de uma vez só, via JavaScript -
    mesmo padrão de `ler_tabela_despesa_certificada` e
    `ler_tabela_domicilio_bancario`.

    Para cada linha clicável (`td.GridLink`) devolve a posição (`indice`,
    usada depois para clicar) e o texto inteiro da linha (`texto`), que
    contém os identificadores da NL e da CE - assim a escolha não depende
    de saber a ordem exata das colunas.
    """
    return popup.locator("#divdtgGerarOrdemCronologica").evaluate(r"""
        div => [...div.querySelectorAll("td.GridLink")].map((td, indice) => ({
            indice: indice,
            texto: (td.closest("tr").innerText || "").replace(/\s+/g, " ").trim()
        }))
    """)


def _indice_registro_ordem_cronologica(popup: "Page", numero_nl: str, numero_ce: str,
                                       numero_linha: int):
    """
    Descobre QUAL linha da grade deve ser clicada: aquela que contém, ao
    mesmo tempo, o número da Nota de Lançamento e o da Despesa
    Certificada da planilha.

    Antes a automação clicava sempre na primeira linha (`.first`), o que
    erra sempre que a grade traz mais de um registro (ou quando ainda
    mostra o resultado da pesquisa anterior).

    A comparação ignora prefixo, pontuação e zeros à esquerda (ex:
    "2026NL065995", "065995" e "65995" são o mesmo número).

    Retorna o índice da linha ou None quando nenhuma linha corresponde
    (nesse caso, registra no log o conteúdo do que o SIGEF mostrou).
    """
    nl_alvo = normalizar_numero_documento(numero_nl)
    ce_alvo = normalizar_numero_documento(numero_ce)

    linhas = ler_tabela_ordem_cronologica(popup)

    if not linhas:
        log_aviso(f"Linha {numero_linha}: a grade da ordem cronológica veio vazia.")
        return None

    # Todos os números de cada linha, já sem zeros à esquerda, para comparar
    # com os identificadores da planilha no mesmo formato.
    numeros_por_linha = [
        (
            registro,
            {str(int(numero)) for numero in REGEX_NUMEROS.findall(registro["texto"])},
        )
        for registro in linhas
    ]

    # 1ª escolha: a linha que traz a NL e a CE juntas (identificação exata).
    for registro, numeros in numeros_por_linha:
        if nl_alvo and ce_alvo and nl_alvo in numeros and ce_alvo in numeros:
            return registro["indice"]

    # 2ª escolha: a linha que traz pelo menos um dos dois identificadores.
    for alvo, rotulo in ((nl_alvo, "NL"), (ce_alvo, "CE")):
        if not alvo:
            continue
        for registro, numeros in numeros_por_linha:
            if alvo in numeros:
                log_info(f"Linha {numero_linha}: registro localizado na grade pela {rotulo} {alvo}.")
                return registro["indice"]

    # Nada reconhecido: com uma única linha, ela é a candidata óbvia; com
    # várias, é melhor pular do que gerar a PP do documento errado.
    if len(linhas) == 1:
        log_aviso(
            f"Linha {numero_linha}: não foi possível reconhecer a NL {nl_alvo} / CE {ce_alvo} "
            f"no texto da grade. Como há um único registro, ele será usado."
        )
        return linhas[0]["indice"]

    log_aviso(
        f"Linha {numero_linha}: a grade trouxe {len(linhas)} registros e nenhum bate com a "
        f"NL {nl_alvo} / CE {ce_alvo}. Pulando para não gerar a PP do documento errado."
    )
    for registro in linhas:
        log_aviso(f"   [{registro['indice']}] {registro['texto']}")

    return None


def _selecionar_registro_ordem_cronologica(page: "Page", popup: "Page", registro_grade,
                                           numero_nl: str, numero_ce: str,
                                           numero_linha: int) -> bool:
    """
    Clica no registro da grade "Gerar Ordem Cronológica" e só devolve o
    controle quando a tela principal confirma que recebeu a Nota de
    Lançamento.

    O SIGEF ocasionalmente ignora o primeiro clique (a janela continua
    aberta e a tela principal segue vazia). Antes, a automação seguia
    assim mesmo e o erro só aparecia lá na frente, na comparação do
    domicílio bancário - que então rodava com os dados da linha anterior.
    Agora o clique é repetido até `TENTATIVAS_CLIQUE_GRADE` vezes e a
    linha só prossegue com a confirmação em mãos.

    Retorna True se o registro foi selecionado; False se a linha deve ser
    pulada.
    """
    for tentativa in range(1, TENTATIVAS_CLIQUE_GRADE + 1):
        try:
            # A linha precisa estar VISÍVEL, e não apenas presente no DOM: a
            # grade do SIGEF já existe (com linhas vazias) antes de os dados
            # chegarem.
            registro_grade.first.wait_for(state="visible", timeout=TIMEOUT_CLIQUE_GRADE)

            indice = _indice_registro_ordem_cronologica(popup, numero_nl, numero_ce, numero_linha)
            if indice is None:
                return False

            registro_grade.nth(indice).click()
        except Exception as erro:
            log_aviso(f"Linha {numero_linha}: falha ao clicar no registro (tentativa {tentativa}): {erro}")

        page.bring_to_front()

        if aguardar_campo_com_valor_por_id(page, "NotaLancamento", TIMEOUT_CLIQUE_GRADE):
            # A tela principal só está pronta para a etapa seguinte quando o
            # postback termina E os campos dela aparecem.
            # Não basta o DOM estar pronto: o SIGEF encadeia postbacks aqui, e
            # agir na janela entre eles fazia a automação perder a seleção do
            # Tipo de Ordem Bancária.
            aguardar_pagina_ociosa(page, timeout=TIMEOUT_CLIQUE_GRADE)

            try:
                page.locator("#cboTipoOrdemBancaria").wait_for(
                    state="visible", timeout=TIMEOUT_CLIQUE_GRADE
                )
                return True
            except Exception:
                log_aviso(
                    f"Linha {numero_linha}: a tela de preparação não terminou de carregar "
                    f"após a seleção (tentativa {tentativa} de {TENTATIVAS_CLIQUE_GRADE})."
                )

        log_aviso(
            f"Linha {numero_linha}: o clique no registro não chegou à tela principal "
            f"(tentativa {tentativa} de {TENTATIVAS_CLIQUE_GRADE}). Repetindo."
        )

        if popup.is_closed():
            break
        popup.bring_to_front()

    log_erro(
        f"Linha {numero_linha}: não foi possível selecionar o registro da ordem "
        f"cronológica. Pulando para a próxima linha."
    )
    return False


def _selecionar_domicilio_bancario(page: "Page", context: "BrowserContext",
                                   banco: str, agencia: str, conta: str) -> bool:
    """
    Abre o popup "Pesquisar Domicílio Bancário", pesquisa pelo banco da
    planilha (coluna E) e seleciona a linha em que os TRÊS dados conferem:
    banco (coluna E), agência (coluna F) e conta (coluna G) - todos
    comparados por `normalizar_chave_bancaria`, nunca como texto bruto
    (a planilha traz "1 / 1178-9 / 74.508-1" e o SIGEF mostra
    "001 / 01178-9 / 0000745081").

    Mesmo padrão de `_selecionar_credor_por_cpf` e
    `_selecionar_empenho_por_numero` (popup -> preencher -> confirmar ->
    selecionar item), inclusive no retorno:

        True  -> domicílio encontrado e selecionado.
        False -> nenhuma linha atende aos três critérios; a linha deve
                 ser pulada e a automação segue para a próxima.
    """
    banco_alvo = normalizar_chave_bancaria(banco)
    agencia_alvo = normalizar_chave_bancaria(agencia)
    conta_alvo = normalizar_chave_bancaria(conta)

    with context.expect_page() as popup_info:
        page.locator("#txtConta_BtnPesquisa").click()

    popup = popup_info.value

    campo_banco = popup.locator("#txtCdBanco")
    campo_banco.wait_for(state="visible")
    campo_banco.fill(banco_alvo.zfill(3))

    mensagem_erro = popup.locator("td.SIGEFMensagemErro")
    resultado_grade = popup.locator("#dtgDomicilioBancario td.GridLink")

    popup.locator("#btnConfirmar").click()

    if not _aguardar_resultado_ou_erro_grade(resultado_grade, mensagem_erro):
        log_aviso(
            f"Banco {banco_alvo}: nenhum domicílio bancário listado pelo SIGEF. "
            f"Pulando para a próxima linha."
        )
        fechar_paginas([popup])
        page.bring_to_front()
        return False

    linhas_grade = ler_tabela_domicilio_bancario(popup)

    encontrado = next(
        (
            linha for linha in linhas_grade
            if normalizar_chave_bancaria(linha["banco"]) == banco_alvo
            and normalizar_chave_bancaria(linha["agencia"]) == agencia_alvo
            and normalizar_chave_bancaria(linha["conta"]) == conta_alvo
        ),
        None,
    )

    if encontrado is None:
        # Loga a chave JÁ NORMALIZADA dos dois lados. Comparar os valores
        # brutos no log engana: caracteres diferentes podem ser idênticos na
        # tela (ver `latinizar_homoglifos`). Com as chaves normalizadas dá
        # para ver na hora QUAL dos três campos não bateu.
        disponiveis = " ; ".join(
            f'{normalizar_chave_bancaria(l["banco"])}/'
            f'{normalizar_chave_bancaria(l["agencia"])}/'
            f'{normalizar_chave_bancaria(l["conta"])}'
            for l in linhas_grade
        ) or "(grade vazia)"

        log_aviso(
            f"Domicílio bancário não encontrado (Banco: {banco} | Agência: {agencia} | "
            f"Conta: {conta}). Pulando para a próxima linha."
        )
        log_info(
            f"    planilha (normalizado): {banco_alvo}/{agencia_alvo}/{conta_alvo}\n"
            f"    SIGEF    (normalizado): {disponiveis}"
        )
        fechar_paginas([popup])
        page.bring_to_front()
        return False

    # O clique fecha o popup. NÃO esperar carregamento nele: a aba está
    # sendo destruída e qualquer `wait_for_load_state` fica pendurado até
    # estourar o timeout - espera-se apenas o evento de fechamento.
    resultado_grade.nth(encontrado["indice"]).click()

    try:
        popup.wait_for_event("close", timeout=TIMEOUT_POPUP_FECHAR)
    except Exception:
        pass  # se não fechou sozinho, `fechar_paginas` resolve

    fechar_paginas([popup])
    page.bring_to_front()
    return True


def _ano_do_documento(codigo_completo: str, ano_padrao=None) -> str:
    """
    Devolve o ano (prefixo "AAAA") de um identificador do SIGEF gravado na
    planilha (ex: "2027NL065995" -> "2027"). Quando a célula traz só o
    número, sem prefixo, usa `ano_padrao` (o ano do exercício da
    automação - ver `constantes.ano_do_exercicio`) ou, sem ele, o ano atual.
    Assim, em janeiro, uma NL "065995" sem prefixo continua sendo procurada
    no exercício certo.
    """
    texto = str(codigo_completo).strip()
    if REGEX_PREFIXO_DOCUMENTO.match(texto):
        return texto[:4]
    return str(ano_padrao or datetime.now().year)


def pp(dados, config=None, worksheet=None):
    """
    Automação: PP (Preparação de Pagamento de Despesa Empenhada).

    Fluxo (repetido por completo a cada linha, já que a confirmação e o
    Limpar resetam a tela a cada iteração):
        1. Conecta 1 única vez à aba do SIGEF na tela de preparação de
           pagamento (a única etapa que ocorre 1 única vez).
        2. Para cada linha da planilha: preenche a gestão e a data de
           referência (Data configurada no sistema) e abre a pesquisa de
           Nota de Lançamento.
        3. Na janela de pesquisa, entra em "Não obedece ordem
           cronológica" e informa a NL (coluna J) e a Despesa Certificada
           (coluna I), ambas sem o prefixo "AAAA+sigla", e seleciona o
           registro.
        4. Seleciona o tipo de ordem bancária e escolhe o domicílio
           bancário cujo banco (coluna E), agência (coluna F) e conta
           (coluna G) conferem com a planilha.
        5. Confirma as retenções e o lançamento. Se o SIGEF devolver
           sucesso, grava o número da PP (com prefixo, ex: "2026PP039772")
           na coluna K; se devolver erro (ex: "Saldo insuficiente..."),
           grava a própria mensagem do SIGEF nessa mesma coluna.
        6. Limpa o formulário (`img[src*='Limpar.GIF']`) e segue para a
           próxima linha.
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return

    if not dados:
        log_aviso("Nenhuma linha para processar na automação PP.")
        return

    config = config or {}
    linha_inicial = config.get("linha_inicial", 2)
    data_referencia = config.get("data", "")

    with sync_playwright() as p:
        try:
            log_info(f"Exercício do SIGEF: {ano_do_exercicio(config)}.")
            context, page = conectar_e_obter_pagina_sigef(p, url_do_exercicio(URL_PP_SIGEF, config), MARCADOR_URL_PP)
        except Exception as erro:
            log_erro(f"Erro ao conectar à tela de preparação de pagamento do SIGEF: {erro}")
            return

        # Locators reaproveitados em TODAS as linhas - buscados uma única
        # vez fora do loop (reduz consultas ao DOM e chamadas ao Playwright).
        campo_gestao = page.locator("#txtGestao_SIGEFPesquisa")
        campo_data_referencia = page.locator("#txtDataReferencia_SIGEFData")
        botao_pesquisar_nl = page.locator("#txtNotaLancamento_BtnPesquisa")
        combo_tipo_ordem = page.locator("#cboTipoOrdemBancaria")

        # Clica em "Limpar" assim que a página carrega, antes de começar o
        # loop - garante que o formulário comece limpo, mesmo que a aba
        # reaproveitada tenha ficado com dados de uma execução anterior.
        try:
            limpar_formulario(page)
        except Exception as erro:
            log_aviso(f"Não foi possível clicar em Limpar ao iniciar a PP: {erro}")

        total_processadas = 0

        # ---- Parte em loop até a última linha da planilha ------------------
        for indice, linha in enumerate(dados):
            if parar_antes_da_linha(linha_inicial + indice):
                break
            with linha_em_execucao("PP", worksheet, linha_inicial + indice, colunas.COL_PP_GERADA, indice + 1, len(dados)) as execucao_linha:
                numero_linha = linha_inicial + indice
                janelas = []

                numero_nl = obter_celula(linha, colunas.COL_PP_NL)
                numero_ce = obter_celula(linha, colunas.COL_PP_CE)
                banco = obter_celula(linha, colunas.COL_PP_BANCO)
                agencia = obter_celula(linha, colunas.COL_PP_AGENCIA)
                conta = obter_celula(linha, colunas.COL_PP_CONTA)

                if not numero_nl or not numero_ce:
                    log_aviso(f"Linha {numero_linha}: NL ou CE vazia. Pulando.")
                    continue

                if not banco or not agencia or not conta:
                    log_aviso(f"Linha {numero_linha}: banco, agência ou conta vazio. Pulando.")
                    continue

                if eh_simulado(numero_nl) or eh_simulado(numero_ce):
                    log_aviso(
                        f"Linha {numero_linha}: a CE ou a NL desta linha foi só SIMULADA - a PP "
                        f"precisa de CE e NL de verdade. Pulando."
                    )
                    continue

                execucao_linha.tentando()
                try:
                    campo_gestao.fill("00001")
                    campo_data_referencia.fill(data_referencia)

                    with context.expect_page() as pagina_pesquisa_info:
                        botao_pesquisar_nl.click()

                    pagina_pesquisa = pagina_pesquisa_info.value
                    janelas.append(pagina_pesquisa)

                    # Espera o link que será usado, em vez de esperar a rede parar.
                    link_fora_ordem = pagina_pesquisa.locator("#lnkNaoObedeceOrdemCronologica")
                    link_fora_ordem.wait_for(state="visible")

                    with context.expect_page() as pagina_lancamento_info:
                        link_fora_ordem.click()

                    pagina_lancamento = pagina_lancamento_info.value
                    janelas.append(pagina_lancamento)

                    campo_sigla_nl = pagina_lancamento.locator("#txtNotaLancamentoSigla")
                    campo_sigla_nl.wait_for(state="visible")

                    ano_exercicio = ano_do_exercicio(config)
                    campo_sigla_nl.fill(_ano_do_documento(numero_nl, ano_exercicio))
                    pagina_lancamento.locator("#txtDespesaCertificadaSigla").fill(_ano_do_documento(numero_ce, ano_exercicio))
                    pagina_lancamento.locator("#txtNotaLancamento_SIGEFPesquisa").fill(
                        extrair_numero_documento(numero_nl).zfill(6)
                    )
                    pagina_lancamento.locator("#txtDespesaCertificada_SIGEFPesquisa").fill(
                        extrair_numero_documento(numero_ce).zfill(6)
                    )
                    pagina_lancamento.locator("#btnConfirmar").click()

                    # Espera o postback da pesquisa terminar e a GRADE existir na
                    # tela antes de qualquer leitura ou clique - sem isso, a
                    # comparação rodava com a tabela ainda vazia.
                    aguardar_pagina_estavel(pagina_lancamento)
                    pagina_lancamento.wait_for_selector(
                        "#divdtgGerarOrdemCronologica", timeout=TIMEOUT_PADRAO_SIGEF
                    )

                    registro_grade = pagina_lancamento.locator("#divdtgGerarOrdemCronologica td.GridLink")
                    mensagem_erro_grade = pagina_lancamento.locator("td.SIGEFMensagemErro")

                    if not _aguardar_resultado_ou_erro_grade(registro_grade, mensagem_erro_grade):
                        log_aviso(
                            f"Linha {numero_linha}: nenhuma preparação encontrada para a NL "
                            f"{numero_nl} / CE {numero_ce}. Pulando para a próxima linha."
                        )
                        fechar_paginas(janelas)
                        page.bring_to_front()
                        continue

                    if not _selecionar_registro_ordem_cronologica(
                        page, pagina_lancamento, registro_grade,
                        numero_nl, numero_ce, numero_linha
                    ):
                        fechar_paginas(janelas)
                        limpar_formulario(page)
                        continue

                    fechar_paginas(janelas)
                    page.bring_to_front()

                    if not selecionar_combo_com_conferencia(page, combo_tipo_ordem, TIPO_ORDEM_BANCARIA_PP):
                        log_erro(
                            f"Linha {numero_linha}: não foi possível selecionar o Tipo de Ordem "
                            f"Bancária. Pulando para a próxima linha."
                        )
                        limpar_formulario(page)
                        continue

                    # Última conferência antes de abrir a pesquisa da conta: se um
                    # postback tardio limpou o combo, a seleção é refeita.
                    if combo_tipo_ordem.input_value() != TIPO_ORDEM_BANCARIA_PP:
                        log_aviso(
                            f"Linha {numero_linha}: o Tipo de Ordem Bancária foi limpo por um "
                            f"recarregamento da tela. Selecionando novamente."
                        )
                        if not selecionar_combo_com_conferencia(page, combo_tipo_ordem, TIPO_ORDEM_BANCARIA_PP):
                            log_erro(
                                f"Linha {numero_linha}: não foi possível manter o Tipo de Ordem "
                                f"Bancária selecionado. Pulando para a próxima linha."
                            )
                            limpar_formulario(page)
                            continue

                    if not _selecionar_domicilio_bancario(page, context, banco, agencia, conta):
                        fechar_paginas(janelas)
                        limpar_formulario(page)
                        continue

                    # MODO SIMULAÇÃO: NL/CE localizadas, tipo de OB e domicílio
                    # bancário conferidos - para ANTES das retenções/confirmação
                    # (que gerariam a PP) e limpa o formulário.
                    if em_simulacao():
                        execucao_linha.simulado_ok(f"NL {numero_nl}, conta {banco}/{agencia}/{conta}")
                        limpar_formulario(page)
                        total_processadas += 1
                        continue

                    # Cada clique espera apenas o elemento SEGUINTE aparecer.
                    clicar_e_aguardar_elemento(page, "#btnRetencoes", "img[src*='aba_confirmacao.gif']")
                    clicar_e_aguardar_elemento(page, "img[src*='aba_confirmacao.gif']", "#btnConfirmar")

                    # Depois deste clique, quem manda é o `wait_for_selector` abaixo.
                    clicar_e_aguardar_elemento(page, "#btnConfirmar")

                    page.wait_for_selector(
                        "td.SIGEFMensagemSucesso, td.SIGEFMensagemErro",
                        timeout=TIMEOUT_PADRAO_SIGEF,
                    )

                    if page.locator("td.SIGEFMensagemErro").count() > 0:
                        # O SIGEF recusou o lançamento (ex: saldo insuficiente ou
                        # data de referência anterior à da NE/NL). O motivo vai
                        # para a coluna K, para aparecer na planilha.
                        mensagem_erro = page.locator("td.SIGEFMensagemErro").first.inner_text().strip()
                        log_erro(f"Linha {numero_linha}: SIGEF recusou a PP - {mensagem_erro}")
                        atualizar_status(worksheet, numero_linha, colunas.COL_PP_GERADA, mensagem_erro)
                        limpar_formulario(page)
                        continue

                    mensagem_sucesso = page.locator("td.SIGEFMensagemSucesso").first.inner_text().strip()
                    busca_pp = REGEX_DOCUMENTO_PP.search(mensagem_sucesso)

                    if busca_pp is None:
                        log_erro(
                            f"Linha {numero_linha}: não foi possível extrair o número da PP "
                            f"da mensagem do SIGEF ({mensagem_sucesso})."
                        )
                        limpar_formulario(page)
                        continue

                    numero_pp = busca_pp.group(0).upper()
                    log_sucesso(f"Linha {numero_linha}: PP '{numero_pp}' gerada.")
                    salvar_valor_gerado(worksheet, numero_linha, colunas.COL_PP_GERADA, numero_pp, rotulo="PP")
                    total_processadas += 1

                    # Limpa o formulário antes de reiniciar o loop na próxima linha.
                    limpar_formulario(page)

                except PlaywrightTimeoutError as erro:
                    log_erro(f"Timeout ao processar a linha {numero_linha} da automação PP: {erro}")
                    fechar_paginas(janelas)
                    page.bring_to_front()
                    try:
                        limpar_formulario(page)
                    except Exception as erro_limpar:
                        log_aviso(f"Linha {numero_linha}: não foi possível clicar em Limpar após o timeout: {erro_limpar}")
                    continue

                except Exception as erro:
                    log_erro(f"Erro ao processar a linha {numero_linha} da automação PP: {erro}")
                    fechar_paginas(janelas)
                    page.bring_to_front()
                    try:
                        limpar_formulario(page)
                    except Exception:
                        pass
                    continue

        log_sucesso(
            f"Automação PP concluída: {total_processadas} de {len(dados)} linha(s) processada(s)."
        )
        return total_processadas


def ler_tabela_pp_geral(page: "Page"):
    """
    Lê de uma vez só a grade `#dtgPP` da tela "Listar Preparação Pagamento
    Geral" - mesmo padrão de `ler_tabela_despesa_certificada` e
    `ler_tabela_domicilio_bancario`: 1 round-trip em vez de 1 chamada ao
    Playwright por célula.

    Devolve, para cada linha clicável (`td.GridLink`):

        indice  - posição na grade, usada depois para clicar sem varrer o
                  DOM de novo.
        texto   - texto da célula clicável, que traz o identificador da PP
                  no formato "AAAAPPnnnnnn" (usado pela Raspar contas).
        celulas - texto de TODAS as células VISÍVEIS da linha (usado pela
                  Raspar PP, que procura o valor da coluna H sem depender
                  da posição fixa da coluna na grade). O filtro
                  `offsetParent !== null` descarta as células ocultas que
                  o SIGEF mantém no HTML.
    """
    return page.locator("#dtgPP").evaluate("""
        table => [...table.querySelectorAll("td.GridLink")].map((td, indice) => ({
            indice: indice,
            texto: (td.textContent || "").trim(),
            celulas: [...td.closest("tr").cells]
                .filter(celula => celula.offsetParent !== null)
                .map(celula => celula.innerText.trim())
        }))
    """)


def _selecionar_favorecido_por_cpf(page: "Page", context: "BrowserContext",
                                   cpf: str) -> bool:
    """
    Seleciona o favorecido pelo CPF na tela "Listar Preparação Pagamento
    Geral".

    NÃO reaproveita `_selecionar_credor_por_cpf` porque esta tela tem um
    passo a mais: o primeiro popup (`#txtFavorecido_BtnPesquisa`) abre uma
    lista genérica, e é o link `#lnkPesquisaCNPJCPFIG` que abre o popup de
    pesquisa por CPF - fechando o primeiro no processo. Por isso o clique
    no link é protegido: o Playwright pode acusar "target closed" quando a
    própria aba que recebeu o clique é destruída, e isso é o comportamento
    esperado, não um erro.

    Retorna:
        True  -> favorecido encontrado e selecionado.
        False -> a pesquisa não retornou nenhuma linha; a linha da
                 planilha deve ser pulada.
    """
    # O campo do SIGEF nesta tela aceita o CPF apenas com dígitos.
    cpf_digitos = REGEX_APENAS_DIGITOS.sub("", str(cpf))

    popup = abrir_popup(context, lambda: page.locator("#txtFavorecido_BtnPesquisa").click())

    with context.expect_page() as popup_cpf_info:
        try:
            popup.locator("#lnkPesquisaCNPJCPFIG").click()
        except PlaywrightError as erro:
            # "closed"/"target closed" aqui é normal: o clique destrói a
            # própria aba. Qualquer outro erro é problema de verdade.
            if "closed" not in str(erro).lower():
                raise

    popup_cpf = popup_cpf_info.value
    aguardar_pagina_estavel(popup_cpf)

    popup_cpf.locator("#btnCPF").click()
    popup_cpf.locator("#txtNuCpf").fill(cpf_digitos)

    mensagem_erro = popup_cpf.locator("td.SIGEFMensagemErro")
    resultado_grade = popup_cpf.locator(
        "#dtgCredor tbody tr:not(.GridCabecalho) td.GridLink"
    )

    popup_cpf.locator("#btnConfirmar").click()

    if not _aguardar_resultado_ou_erro_grade(resultado_grade, mensagem_erro):
        log_aviso(
            f"CPF {cpf_digitos}: nenhum favorecido encontrado no SIGEF. "
            f"Pulando para a próxima linha."
        )
        fechar_paginas([popup_cpf, popup])
        page.bring_to_front()
        return False

    resultado_grade.first.click()
    fechar_paginas([popup_cpf, popup])
    page.bring_to_front()
    return True


def _valor_planilha_para_float(valor) -> float:
    """
    Converte o valor da planilha (coluna H) para float em reais.

    A planilha guarda "840" querendo dizer R$ 840,00, enquanto a grade do
    SIGEF exibe "840,00" - sem essa conversão a comparação nunca bateria.
    Reaproveita `formatar_valor_centavos`, que já é a regra oficial do
    sistema para ler esse formato (e também aceita "1.400,00" e "1400,00").
    """
    return int(formatar_valor_centavos(valor)) / 100


def _linha_grade_tem_valor(celulas, valor_esperado: float) -> bool:
    """
    Diz se alguma célula VISÍVEL da linha da grade é o valor esperado.

    Procura em TODAS as células em vez de olhar uma posição fixa: assim a
    automação não quebra se o SIGEF mudar a ordem das colunas da grade.
    Só tenta converter as células que têm cara de valor monetário BR
    (`REGEX_VALOR_BR`), e a comparação usa tolerância de meio centavo,
    porque float não é exato.
    """
    for celula in celulas:
        texto = str(celula).replace("R$", "").strip()
        if REGEX_VALOR_BR.match(texto):
            if abs(_valor_br_para_float(texto) - valor_esperado) < 0.005:
                return True
    return False


def raspar_pp(dados, config=None, worksheet=None):
    """
    Automação: Raspar PP (Listar Preparação Pagamento Geral).

    Recupera a PP que já foi gerada no SIGEF para cada linha da planilha -
    útil quando a automação PP rodou mas o número não chegou a ser gravado
    (queda de conexão, execução manual etc.).

    Fluxo:
        1. Conecta 1 única vez à tela "Listar Preparação Pagamento Geral".
        2. Para cada linha da planilha:
           a. Preenche a parte fixa (unidade gestora 160001 / gestão 00001).
           b. Pesquisa e seleciona o favorecido pelo CPF (coluna B) e
              confirma - o SIGEF lista TODAS as PPs daquele favorecido.
           c. Procura na grade a(s) linha(s) cujo valor bate com a coluna H.
              Havendo mais de uma, usa a ÚLTIMA (a mais recente).
           d. Abre o detalhe e captura CE, NL e PP.
           e. Se CE e NL conferem com as colunas I e J, grava a PP na
              coluna K e a mensagem de conferência na coluna M. Se NÃO
              conferem, NÃO grava a PP - apenas registra a divergência.
        3. Erro em uma linha nunca interrompe a automação - registra e
           segue para a próxima.

    Toda comparação é feita sobre valores normalizados, nunca texto bruto:
    números de documento por `normalizar_numero_documento` (que resolve o
    "18487" da planilha contra o "2026CE018487" do SIGEF) e valores por
    `_valor_planilha_para_float` (que resolve o "840" contra "840,00").
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return

    if not dados:
        log_aviso("Nenhuma linha para processar na automação Raspar PP.")
        return

    config = config or {}
    linha_inicial = config.get("linha_inicial", 2)

    with sync_playwright() as p:
        try:
            context, page = conectar_e_obter_pagina_sigef(
                p, url_do_exercicio(URL_RASPAR_CONTA_SIGEF, config), MARCADOR_URL_RASPAR_CONTA
            )
        except Exception as erro:
            log_erro(f"Erro ao conectar à tela de listagem de PP do SIGEF: {erro}")
            return

        campo_unidade_gestora = page.locator("#txtUnidadeGestora")
        campo_gestao = page.locator("#txtGestao_SIGEFPesquisa")
        botao_confirmar = page.locator("#btnConfirmar")
        linhas_grid_link = page.locator("#dtgPP td.GridLink")

        total_processadas = 0
        total_correspondentes = 0
        total_divergentes = 0

        for indice, linha in enumerate(dados):
            if parar_antes_da_linha(linha_inicial + indice):
                break
            informar_progresso(indice + 1, len(dados))
            numero_linha = linha_inicial + indice

            cpf = obter_celula(linha, colunas.COL_RASPAR_PP_CPF)
            valor_bruto = obter_celula(linha, colunas.COL_RASPAR_PP_VALOR)
            ce_esperada = obter_celula(linha, colunas.COL_RASPAR_PP_CE)
            nl_esperada = obter_celula(linha, colunas.COL_RASPAR_PP_NL)

            if not cpf or not valor_bruto:
                log_aviso(f"Linha {numero_linha}: CPF ou valor (coluna H) vazio. Pulando.")
                continue

            if eh_simulado(ce_esperada) or eh_simulado(nl_esperada):
                log_aviso(f"Linha {numero_linha}: CE ou NL só simulada - nada para buscar. Pulando.")
                continue

            if not ce_esperada or not nl_esperada:
                log_aviso(
                    f"Linha {numero_linha}: CE (coluna I) ou NL (coluna J) vazia - "
                    f"sem elas não há como conferir a PP. Pulando."
                )
                continue

            try:
                valor_esperado = _valor_planilha_para_float(valor_bruto)
            except (ValueError, TypeError):
                log_aviso(
                    f"Linha {numero_linha}: valor '{valor_bruto}' (coluna H) não é "
                    f"numérico. Pulando."
                )
                continue

            log_info(
                f"Linha {numero_linha}: procurando PP de {formatar_valor_br(valor_bruto)} "
                f"para o CPF {formatar_cpf(cpf)}..."
            )

            try:
                # Parte fixa refeita A CADA linha: a tela faz postback ao
                # confirmar e ao selecionar o favorecido, e os campos podem
                # voltar em branco.
                campo_unidade_gestora.fill("160001")
                campo_gestao.fill("00001")

                if not _selecionar_favorecido_por_cpf(page, context, cpf):
                    continue

                botao_confirmar.click()
                aguardar_pagina_estavel(page)

                registros = ler_tabela_pp_geral(page)
                if not registros:
                    log_aviso(
                        f"Linha {numero_linha}: nenhuma PP listada para este favorecido. "
                        f"Pulando para a próxima linha."
                    )
                    continue

                candidatos = [
                    registro for registro in registros
                    if len(registro["celulas"]) > 1
                    and _linha_grade_tem_valor(registro["celulas"], valor_esperado)
                ]

                if not candidatos:
                    log_aviso(
                        f"Linha {numero_linha}: nenhuma PP na grade com o valor "
                        f"{formatar_valor_br(valor_bruto)}. Pulando para a próxima linha."
                    )
                    continue

                # Havendo mais de uma PP com o mesmo valor, a ÚLTIMA da grade
                # é a mais recente - é ela que interessa.
                escolhido = candidatos[-1]
                if len(candidatos) > 1:
                    log_aviso(
                        f"Linha {numero_linha}: {len(candidatos)} PPs com o valor "
                        f"{formatar_valor_br(valor_bruto)}. Usando a última (mais recente)."
                    )

                detalhe = abrir_popup(context, lambda: linhas_grid_link.nth(escolhido["indice"]).click())

                ce_sigef = detalhe.locator("#txtNuDespesaCertificada").input_value()
                nl_sigef = detalhe.locator("#txtNotaLancamento").input_value()
                pp_sigef = detalhe.locator("#txtNuPreparacaoPagamento").input_value()

                try:
                    detalhe.locator("#SIGEFBotoesImpressao_BtnFechar").click()
                except Exception:
                    # O botão pode não estar presente/visível; `fechar_paginas`
                    # encerra a aba de qualquer forma.
                    pass
                fechar_paginas([detalhe])
                page.bring_to_front()

                ce_confere = (
                    normalizar_numero_documento(ce_sigef)
                    == normalizar_numero_documento(ce_esperada)
                )
                nl_confere = (
                    normalizar_numero_documento(nl_sigef)
                    == normalizar_numero_documento(nl_esperada)
                )

                if ce_confere and nl_confere:
                    log_sucesso(
                        f"Linha {numero_linha}: {pp_sigef} corresponde à CE e NL da planilha."
                    )
                    salvar_valor_gerado(
                        worksheet, numero_linha, colunas.COL_PP_GERADA, pp_sigef, rotulo="Raspar PP"
                    )
                    atualizar_status(
                        worksheet, numero_linha, colunas.COL_RASPAR_PP_OBSERVACAO,
                        f"PP {pp_sigef} correspondente à CE {ce_sigef} e NL {nl_sigef}"
                    )
                    total_correspondentes += 1
                else:
                    # PP divergente NÃO vai para a coluna K - gravar um número
                    # errado ali contaminaria as automações seguintes, que
                    # leem essa coluna.
                    divergencias = []
                    if not ce_confere:
                        divergencias.append(f"CE {ce_sigef} != {ce_esperada}")
                    if not nl_confere:
                        divergencias.append(f"NL {nl_sigef} != {nl_esperada}")

                    log_aviso(
                        f"Linha {numero_linha}: {pp_sigef} NÃO corresponde à CE e NL da "
                        f"planilha ({'; '.join(divergencias)}). PP não gravada na coluna K."
                    )
                    atualizar_status(
                        worksheet, numero_linha, colunas.COL_RASPAR_PP_OBSERVACAO,
                        f"PP {pp_sigef} NÃO corresponde: {'; '.join(divergencias)}"
                    )
                    total_divergentes += 1

                total_processadas += 1

            except PlaywrightTimeoutError as erro:
                log_erro(f"Timeout ao processar a linha {numero_linha} da Raspar PP: {erro}")
                continue
            except Exception as erro:
                log_erro(f"Erro ao processar a linha {numero_linha} da Raspar PP: {erro}")
                continue

        log_sucesso(
            f"Automação Raspar PP concluída: {total_processadas} de {len(dados)} "
            f"linha(s) processada(s) - {total_correspondentes} correspondente(s) e "
            f"{total_divergentes} divergente(s)."
        )
        return total_processadas


