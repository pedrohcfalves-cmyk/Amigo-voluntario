"""Automação NL (Nota de Lançamento / Liquidação) e Raspar NL."""
import time

from .. import colunas
from ..constantes import (
    URL_RASPAR_CE_SIGEF, MARCADOR_URL_RASPAR_CE, URL_NL_SIGEF, MARCADOR_URL_NL,
    TIMEOUT_PADRAO_SIGEF,
)
from ..excel import obter_celula, salvar_valor_gerado
from ..log import log_info, log_sucesso, log_erro, log_aviso
from ..navegador import (
    conectar_e_obter_pagina_sigef, aguardar_pagina_estavel, limpar_formulario,
    aguardar_campo_preenchido, abrir_popup,
)
from ..playwright_compat import sync_playwright, Page, PlaywrightTimeoutError
from ..utils import (
    extrair_numero_documento, normalizar_numero_documento, formatar_valor_centavos,
    formatar_valor_br,
)
from .ce import _selecionar_empenho_por_numero

def _extrair_documentos_grade(page: "Page"):
    """
    Extrai todos os pares (documento, valor_líquido) das linhas
    preenchidas da grade `#dtgDocumentos` - o SIGEF pode devolver mais de
    uma NL para uma única liquidação (ex: quando alguma retenção obriga a
    dividir em mais de um documento).

    Usa reconhecimento por PADRÃO em vez de índice fixo de coluna
    (`.nth(5)`/`.nth(7)` se mostraram frágeis, pois a grade pode variar
    de estrutura entre execuções): identifica a célula do documento pelo
    formato "AAAA" + 2 letras + dígitos (ex: "2026NL066020") e assume o
    valor líquido como a ÚLTIMA célula alinhada à direita da linha (o
    Valor Líquido sempre vem depois do Valor Bruto) - mesma estratégia já
    usada com sucesso em `_extrair_ultima_despesa_certificada` (Raspar CE).
    """
    linhas = page.locator("#dtgDocumentos").evaluate("""
        table => {
            const padraoDocumento = /^\\d{4}[A-Z]{2}\\d+$/;

            return [...table.querySelectorAll("tr")]
                .map(tr => {
                    const tds = [...tr.querySelectorAll("td")];
                    const tdDocumento = tds.find(td => padraoDocumento.test(td.textContent.trim()));
                    if (!tdDocumento) return null;

                    const valoresDireita = tds
                        .filter(td => td.align === "right")
                        .map(td => td.textContent.trim());

                    return {
                        documento: tdDocumento.textContent.trim(),
                        valor_liquido: valoresDireita.length ? valoresDireita[valoresDireita.length - 1] : ""
                    };
                })
                .filter(item => item !== null);
        }
    """)

    return [(item["documento"], item["valor_liquido"]) for item in linhas if item["valor_liquido"]]


def _valor_br_para_float(texto: str) -> float:
    """Converte um valor exibido no padrão BR pelo SIGEF (ex: '1.260,00')
    de volta para float, para permitir somar valores de múltiplos
    documentos antes de comparar com o esperado pela planilha."""
    texto = str(texto).replace("R$", "").strip()
    return float(texto.replace(".", "").replace(",", ".")) if texto else 0.0


def _selecionar_documento_correspondente(documentos, valor_liquido_esperado: str):
    """
    Decide qual(is) documento(s) de uma lista (documento, valor_líquido)
    correspondem ao valor esperado pela planilha, com a mesma prioridade
    usada tanto na checagem prévia (documento já lançado antes de
    liquidar de novo) quanto na checagem final (após confirmar):

        1. Algum documento, sozinho, já bate exatamente com o esperado -
           usa só esse (mesmo que existam outras linhas na grade).
        2. Só existe 1 documento na grade - usa ele.
        3. Nenhum bate sozinho - soma os valores líquidos de todos e
           combina os números na mesma célula.

    Retorna (numero_documento, valor_liquido) - `numero_documento` já vem
    combinado com " + " quando forem 2 ou mais somados.
    """
    correspondencia_unica = next(
        (doc for doc in documentos if doc[1] == valor_liquido_esperado), None
    )
    if correspondencia_unica is not None:
        return correspondencia_unica

    if len(documentos) == 1:
        return documentos[0]

    soma = sum(_valor_br_para_float(valor) for _, valor in documentos)
    numero_combinado = " + ".join(documento for documento, _ in documentos)
    return numero_combinado, formatar_valor_br(soma)


def nl(dados, config=None, worksheet=None):
    """
    Automação: NL (Liquidar Despesa Certificada).

    Fluxo (repetido por completo a cada linha, já que a pesquisa por CE e
    o preenchimento do formulário resetam a tela a cada iteração):
        1. Conecta 1 única vez à aba do SIGEF na tela de liquidação
           (a única etapa que realmente ocorre 1 única vez).
        2. Para cada linha da planilha: pesquisa a despesa certificada
           pelo número da CE (coluna I). Se a grade já mostrar alguma NL
           lançada e o valor líquido bater com o esperado (coluna H),
           salva direto na coluna J, avisa e reinicia o ciclo - sem
           lançar uma nova liquidação por cima. Caso contrário, prossegue
           normalmente: define a data de vencimento (Data configurada no
           sistema) e adiciona a liquidação.
        3. Pesquisa e seleciona a Nota de Empenho vinculada (coluna D).
        4. Preenche o valor bruto (coluna H, convertido para centavos) e
           confirma as retenções.
        5. Extrai o(s) documento(s) e valor(es) líquido(s) gerados pelo
           SIGEF na grade. Se algum documento, sozinho, já bater
           exatamente com o valor esperado, usa só esse (mesmo que
           existam outras linhas na grade). Caso contrário, e se vierem
           2 ou mais documentos, soma os valores líquidos de todos e
           compara a soma com o esperado. Em qualquer um dos casos, se
           conferir, grava o(s) número(s) da NL na coluna J (quando forem
           2+ combinados na soma, todos juntos na MESMA célula).
        6. Limpa o formulário (`img[src*='Limpar.GIF']`) e segue para a
           próxima linha.
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return

    if not dados:
        log_aviso("Nenhuma linha para processar na automação NL.")
        return

    config = config or {}
    linha_inicial = config.get("linha_inicial", 2)
    data_vencimento = config.get("data", "")

    with sync_playwright() as p:
        try:
            context, page = conectar_e_obter_pagina_sigef(p, URL_NL_SIGEF, MARCADOR_URL_NL)
        except Exception as erro:
            log_erro(f"Erro ao conectar à tela de liquidação de despesa certificada do SIGEF: {erro}")
            return

        # Locators reaproveitados em TODAS as linhas - buscados uma única
        # vez fora do loop (reduz consultas ao DOM e chamadas ao Playwright).
        campo_gestao = page.locator("#txtCdGestao_SIGEFPesquisa")
        campo_ce_pesquisa = page.locator("#txtDespesaCertificadaNumero_SIGEFPesquisa")
        botao_pesquisar = page.locator("#btnPesquisar")
        campo_data_vencimento = page.locator("#txtDataVencimento_SIGEFData")
        botao_adicionar = page.locator("#btnAdicionar")
        campo_valor_bruto = page.locator("#txtValorBrutoId")
        botao_retencoes = page.locator("#btnRetencoesId")
        botao_menu_n4 = page.locator("#menun4")
        botao_confirmar = page.locator("#btnConfirmar")

        # Clica em "Limpar" assim que a página carrega, antes de começar o
        # loop - garante que o formulário comece limpo, mesmo que a aba
        # reaproveitada tenha ficado com dados de uma execução anterior.
        try:
            limpar_formulario(page)
        except Exception as erro:
            log_aviso(f"Não foi possível clicar em Limpar ao iniciar a NL: {erro}")

        total_processadas = 0

        # ---- Parte em loop até a última linha da planilha ------------------
        for indice, linha in enumerate(dados):
            numero_linha = linha_inicial + indice

            numero_ce = obter_celula(linha, colunas.COL_NL_CE)
            numero_ne = obter_celula(linha, colunas.COL_NL_NE)
            valor_bruto = obter_celula(linha, colunas.COL_CE_VALOR)

            if not numero_ce or not numero_ne or not valor_bruto:
                log_aviso(f"Linha {numero_linha}: CE, NE ou valor vazio. Pulando.")
                continue

            try:
                campo_gestao.fill("00001")
                campo_ce_pesquisa.fill(extrair_numero_documento(numero_ce).zfill(6))
                botao_pesquisar.click()
                aguardar_pagina_estavel(page)

                # Verifica se essa despesa certificada JÁ TEM alguma NL
                # lançada antes de prosseguir - se já existir e o valor
                # bater (mesmo critério da checagem final, após
                # confirmar), só salva e reinicia o ciclo sem lançar uma
                # nova liquidação por cima.
                valor_liquido_esperado = formatar_valor_br(valor_bruto)
                documentos_existentes = _extrair_documentos_grade(page)

                if documentos_existentes:
                    nl_existente, valor_liquido_existente = _selecionar_documento_correspondente(
                        documentos_existentes, valor_liquido_esperado
                    )

                    if valor_liquido_existente == valor_liquido_esperado:
                        log_sucesso(
                            f"Linha {numero_linha}: NL '{nl_existente}' já lançada, valor líquido "
                            f"{valor_liquido_existente} confere com a planilha. NL salva."
                        )
                        salvar_valor_gerado(worksheet, numero_linha, colunas.COL_NL_GERADA, nl_existente, rotulo="NL")
                        limpar_formulario(page)
                        total_processadas += 1
                        continue

                    log_aviso(
                        f"Linha {numero_linha}: já existe(m) documento(s) na grade, mas o valor "
                        f"não confere ({valor_liquido_existente} != {valor_liquido_esperado}). "
                        f"Prosseguindo com o lançamento normal."
                    )

                campo_data_vencimento.fill(data_vencimento)
                botao_adicionar.click()

                if not _selecionar_empenho_por_numero(page, context, extrair_numero_documento(numero_ne)):
                    continue

                # Requisito obrigatório: manter `press_sequentially()`
                # exatamente como está - não substituir por `.fill()`.
                valor_centavos = formatar_valor_centavos(valor_bruto)
                campo_valor_bruto.press_sequentially(valor_centavos)

                # Só prossegue depois que o campo confirma ter recebido TODO o
                # valor - a digitação tecla a tecla pode ficar para trás e o
                # clique seguinte acontecer com o valor pela metade.
                if not aguardar_campo_preenchido(campo_valor_bruto, valor_centavos):
                    log_erro(
                        f"Linha {numero_linha}: o valor bruto não foi digitado por completo "
                        f"no SIGEF (esperado {valor_centavos}, campo com "
                        f"'{campo_valor_bruto.input_value()}'). Pulando para a próxima linha."
                    )
                    limpar_formulario(page)
                    continue

                botao_retencoes.click()
                botao_menu_n4.click()
                botao_confirmar.click()

                # Aguarda a grade recarregar (postback) fazendo polling da
                # PRÓPRIA extração, em vez de esperar a linha ficar
                # "visível" - as linhas (mesmo vazias, só com "&nbsp;") já
                # ficam visíveis antes do SIGEF preencher os dados de
                # verdade, então esperar só por visibilidade resolvia cedo
                # demais e a extração rodava com a grade ainda vazia.
                aguardar_pagina_estavel(page)

                documentos = []
                prazo = time.monotonic() + (TIMEOUT_PADRAO_SIGEF / 1000)
                while time.monotonic() < prazo:
                    documentos = _extrair_documentos_grade(page)
                    if documentos:
                        break
                    time.sleep(0.15)

                if not documentos:
                    total_tabelas = page.locator("#dtgDocumentos").count()
                    total_linhas = page.locator("#dtgDocumentos tr").count()
                    log_erro(
                        f"Linha {numero_linha}: nenhum documento encontrado na grade após "
                        f"confirmar (tabela #dtgDocumentos encontrada: {total_tabelas}x, "
                        f"total de linhas <tr> dentro dela: {total_linhas})."
                    )
                    continue

                # valor_liquido_esperado já foi calculado antes do
                # Pesquisar (reaproveitado aqui, mesmo valor da linha).
                nl_gerada, valor_liquido_sigef = _selecionar_documento_correspondente(
                    documentos, valor_liquido_esperado
                )

                if valor_liquido_sigef == valor_liquido_esperado:
                    log_sucesso(
                        f"Linha {numero_linha}: NL '{nl_gerada}' gerada, valor líquido "
                        f"{valor_liquido_sigef} confere com a planilha."
                    )
                    salvar_valor_gerado(worksheet, numero_linha, colunas.COL_NL_GERADA, nl_gerada, rotulo="NL")
                    total_processadas += 1
                else:
                    log_erro(
                        f"Linha {numero_linha}: valor líquido do SIGEF ({valor_liquido_sigef}) "
                        f"não confere com o esperado pela planilha ({valor_liquido_esperado}). NL não salva."
                    )

                # Limpa o formulário antes de reiniciar o loop na próxima linha.
                limpar_formulario(page)

            except PlaywrightTimeoutError as erro:
                log_erro(f"Timeout ao processar a linha {numero_linha} da automação NL: {erro}")
                try:
                    limpar_formulario(page)
                except Exception as erro_limpar:
                    log_aviso(f"Linha {numero_linha}: não foi possível clicar em Limpar após o timeout: {erro_limpar}")
                continue

            except Exception as erro:
                log_erro(f"Erro ao processar a linha {numero_linha} da automação NL: {erro}")
                continue

        log_sucesso(
            f"Automação NL concluída: {total_processadas} de {len(dados)} linha(s) processada(s)."
        )


def ler_tabela_despesa_certificada(page: "Page"):
    """
    Lê TODA a grade `#dtgDespesaCertificada` de uma única vez, via
    JavaScript (`evaluate`) - evita dezenas/centenas de chamadas ao
    Playwright (1 locator + 1 round-trip, em vez de 1 chamada por
    célula). Para cada linha com um registro clicável (`td.GridLink`),
    captura Número, Valor, Documento e Favorecido - preparado para
    futuras expansões (basta acrescentar mais campos ao objeto JS
    retornado).

    Offsets confirmados a partir da célula do Número (`td.GridLink`):
    1 célula depois = Tipo, 2 células depois = Número Documento, 3
    células depois = Favorecido (o texto visível dessa célula é o CPF; o
    nome completo do favorecido vem no atributo `title`).
    """
    return page.locator("#dtgDespesaCertificada").evaluate("""
        table => {
            const linhas = [...table.querySelectorAll("tr")]
                .filter(tr => tr.querySelector("td.GridLink"));

            return linhas.map(tr => {
                const tdNumero = tr.querySelector("td.GridLink");

                const valoresDireita = [...tr.querySelectorAll("td")]
                    .filter(td => td.align === "right")
                    .map(td => td.textContent.trim());

                const tdTipo = tdNumero.nextElementSibling;
                const tdDocumento = tdTipo ? tdTipo.nextElementSibling : null;
                const tdFavorecido = tdDocumento ? tdDocumento.nextElementSibling : null;

                const favorecido = tdFavorecido
                    ? (tdFavorecido.getAttribute("title") || tdFavorecido.textContent.trim())
                    : "";

                return {
                    numero: tdNumero.textContent.trim(),
                    documento: tdDocumento ? tdDocumento.textContent.trim() : "",
                    favorecido: favorecido,
                    valor: valoresDireita.length ? valoresDireita[valoresDireita.length - 1] : ""
                };
            });
        }
    """)


def raspar_nl(dados, config=None, worksheet=None):
    """
    Automação: Raspar NL (Listar Despesa Certificada Geral - pesquisa
    direta pelo número da despesa).

    Fluxo:
        1. Conecta 1 única vez à tela de listagem de despesa certificada
           (mesma tela/URL da Raspar CE).
        2. Parte fixa (1 única vez): clica em Limpar, preenche unidade
           gestora e gestão.
        3. Para cada linha da planilha: pesquisa pelo número da CE
           (coluna I, já com zeros à esquerda), lê a tabela inteira de
           uma vez (`ler_tabela_despesa_certificada`), normaliza os
           números (remove prefixo "AAAACE" e zeros à esquerda) e compara
           com o número esperado da planilha.
        4. Se encontrar: abre os detalhes (despesa -> Nota de Lançamento
           -> popup), captura CE/NL/NE e salva a NL na coluna J.
        5. Se NÃO encontrar (ou a tabela vier vazia): registra o aviso e
           segue para a próxima linha, sem interromper a automação.
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return

    if not dados:
        log_aviso("Nenhuma linha para processar na automação Raspar NL.")
        return

    config = config or {}
    linha_inicial = config.get("linha_inicial", 2)

    with sync_playwright() as p:
        try:
            context, page = conectar_e_obter_pagina_sigef(p, URL_RASPAR_CE_SIGEF, MARCADOR_URL_RASPAR_CE)
        except Exception as erro:
            log_erro(f"Erro ao conectar à tela de listagem de despesa certificada do SIGEF: {erro}")
            return

        # ---- Parte fixa: executa 1 única vez -------------------------------
        try:
            page.locator("#btnBotoesImpressao_BtnLimpar").click(timeout=5000)
        except Exception as erro:
            log_aviso(
                f"Não foi possível clicar em Limpar ao iniciar a Raspar NL "
                f"(botão pode não existir nesta tela): {erro}"
            )

        page.locator("#txtCdUnidadeGestora").fill("160001")
        page.locator("#txtCdGestao_SIGEFPesquisa").fill("00001")

        # Locators reaproveitados em TODAS as linhas - buscados uma única
        # vez fora do loop (reduz consultas ao DOM e chamadas ao Playwright).
        campo_despesa = page.locator("#txtNuDespesa")
        botao_confirmar = page.locator("#btnConfirmar")
        linhas_grid_link = page.locator("#dtgDespesaCertificada td.GridLink")

        total_processadas = 0

        # ---- Parte em loop até a última linha da planilha ------------------
        for indice, linha in enumerate(dados):
            numero_linha = linha_inicial + indice

            numero_ce_bruto = obter_celula(linha, colunas.COL_NL_CE)
            if not numero_ce_bruto:
                log_aviso(f"Linha {numero_linha}: CE vazia. Pulando.")
                continue

            numero_ce_normalizado = normalizar_numero_documento(numero_ce_bruto)
            numero_pesquisa = extrair_numero_documento(numero_ce_bruto).zfill(6)

            log_info(f"Pesquisando CE/NL {numero_ce_normalizado}...")

            try:
                campo_despesa.fill(numero_pesquisa)
                botao_confirmar.click()
                aguardar_pagina_estavel(page)

                registros = ler_tabela_despesa_certificada(page)
                log_info("Tabela carregada.")

                if not registros:
                    log_aviso(
                        f"Linha {numero_linha}: tabela veio vazia para a CE/NL "
                        f"{numero_ce_normalizado}. Pulando para a próxima linha."
                    )
                    continue

                log_info("Comparando registros...")
                indice_encontrado = next(
                    (
                        i for i, registro in enumerate(registros)
                        if normalizar_numero_documento(registro["numero"]) == numero_ce_normalizado
                    ),
                    None,
                )

                if indice_encontrado is None:
                    log_aviso(
                        f"Linha {numero_linha}: nenhum registro correspondente à CE/NL "
                        f"{numero_ce_normalizado} encontrado. Pulando para a próxima linha."
                    )
                    continue

                log_sucesso("Registro encontrado.")
                log_info("Abrindo detalhes...")

                nova_pagina = abrir_popup(context, lambda: linhas_grid_link.nth(indice_encontrado).click())
                nova_pagina.locator("#btnNotaLancamento").click()

                popup_nl = nova_pagina.locator("#dtgNotaLancamento td.GridLink")
                if popup_nl.count() == 0:
                    log_aviso(
                        f"Linha {numero_linha}: nenhuma Nota de Lançamento encontrada para "
                        f"a CE {numero_ce_normalizado}. Pulando para a próxima linha."
                    )
                    nova_pagina.close()
                    page.bring_to_front()
                    continue

                pop_up = abrir_popup(context, lambda: popup_nl.first.click())

                ce_capturada = pop_up.locator("#txtDespesaCeritificada").input_value()
                nl_capturada = pop_up.locator("#txtDocumentoOriginal").input_value()
                ne_capturada = pop_up.locator("#txtNotaEmpenhoOriginal").input_value()

                log_sucesso(f"Linha {numero_linha}: CE {ce_capturada} | NL {nl_capturada} | NE {ne_capturada}.")

                # Salva a NL apenas na coluna J (mesma coluna usada pela
                # automação NL).
                salvar_valor_gerado(worksheet, numero_linha, colunas.COL_NL_GERADA, nl_capturada, rotulo="Raspar NL")

                pop_up.locator("#btnImpressao_BtnFechar").click()
                pop_up.close()
                nova_pagina.close()
                page.bring_to_front()

                total_processadas += 1

            except PlaywrightTimeoutError as erro:
                log_erro(f"Timeout ao processar a linha {numero_linha} da automação Raspar NL: {erro}")
                continue
            except Exception as erro:
                log_erro(f"Erro ao processar a linha {numero_linha} da automação Raspar NL: {erro}")
                continue

        log_sucesso(
            f"Automação Raspar NL concluída: {total_processadas} de {len(dados)} linha(s) processada(s)."
        )


