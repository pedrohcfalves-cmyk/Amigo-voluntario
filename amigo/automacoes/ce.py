"""Automação CE (Despesa Certificada) e Raspar CE."""
import time

from .. import colunas
from ..constantes import (
    URL_CE_SIGEF, MARCADOR_URL_CE, URL_RASPAR_CE_SIGEF, MARCADOR_URL_RASPAR_CE,
    TIMEOUT_PADRAO_SIGEF,
)
from ..excel import obter_celula, salvar_valor_gerado
from ..log import log_info, log_sucesso, log_erro, log_aviso
from ..navegador import (
    conectar_e_obter_pagina_sigef, aguardar_pagina_estavel, aguardar_campo_com_valor_por_id,
    abrir_popup,
)
from ..playwright_compat import sync_playwright, Page, BrowserContext
from ..utils import formatar_cpf, formatar_valor_centavos, obter_numero_mes_referencia

def _preencher_campos_fixos_ce(page: "Page", config: dict) -> None:
    """
    Preenche os campos fixos da tela "Manter Despesa Certificada" (CE) -
    os MESMOS para todas as linhas da planilha, portanto executados uma
    única vez por automação (parte fixa).
    """
    mes_referencia = config.get("mes_referencia", "")
    processo = config.get("processo", "")
    data = config.get("data", "")

    page.locator("#txtCdGestao_SIGEFPesquisa").fill("00001")
    page.locator("#txtNuDocumento").fill(mes_referencia)
    page.locator("#chkFlAtestadoRecSouResp").check()
    page.locator("#txtDeObservacao").fill(
        f"Gratificação amigo voluntario {mes_referencia} Processo: {processo}"
    )
    page.locator("#txtDtAceite_SIGEFData").fill(data)
    page.locator("#txtDtApresentacao_SIGEFData").fill(data)
    page.locator("#txtDtEmissao_SIGEFData").fill(data)
    page.locator("#cboMesComp").select_option(obter_numero_mes_referencia(mes_referencia))


def _aguardar_resultado_ou_erro_grade(resultado_grade, mensagem_erro,
                                       timeout_ms: int = TIMEOUT_PADRAO_SIGEF) -> bool:
    """
    Faz polling curto (a cada 150ms) esperando QUALQUER UM dos dois
    desfechos possíveis de uma pesquisa em popup do SIGEF (credor, Nota
    de Empenho etc.): uma linha aparecer na grade OU a mensagem "Não há
    registros a serem listados." ficar visível - respondendo assim que
    qualquer um dos dois ocorrer, em vez de esperar o timeout inteiro
    sempre que o SIGEF já respondeu rapidamente com "sem resultados" (era
    o que causava a demora extra para pular de linha).

    Retorna True se encontrou ao menos 1 linha na grade, False se a
    mensagem de erro apareceu (e permaneceu) ou se o timeout total
    expirar sem nenhum dos dois sinais.
    """
    prazo = time.monotonic() + (timeout_ms / 1000)

    while time.monotonic() < prazo:
        if resultado_grade.count() > 0:
            return True

        if mensagem_erro.first.is_visible():
            # A mensagem de erro pode aparecer BREVEMENTE antes da grade
            # terminar de renderizar (falso positivo transitório - foi o
            # caso do CPF 764.167.802-59, que aparecia normalmente como
            # 1ª linha da grade, mas era marcado como "não encontrado").
            # Por isso, espera um instante e reconfirma os dois sinais
            # antes de aceitar a mensagem de erro como definitiva.
            time.sleep(0.3)
            if resultado_grade.count() > 0:
                return True
            if mensagem_erro.first.is_visible():
                return False
        time.sleep(0.15)

    return resultado_grade.count() > 0


def _selecionar_credor_por_cpf(
    page: "Page",
    context: "BrowserContext",
    cpf_formatado: str,
    seletor_botao_pesquisa: str = "#txtNmCredor_BtnPesquisa",
) -> bool:
    """
    Abre o popup de pesquisa de credor, pesquisa pelo CPF informado e
    seleciona o primeiro resultado da grade - reproduz o fluxo manual
    (clicar em pesquisar -> CPF -> confirmar -> selecionar item) e traz o
    foco de volta para a página principal ao final.

    Compartilhada entre as automações CE (`#txtNmCredor_BtnPesquisa`) e
    Raspar CE (`#txtCdCredor_BtnPesquisa`) - o botão de pesquisa muda de
    uma tela para outra, mas o restante do fluxo do popup é idêntico.

    Retorna:
        True  -> credor encontrado e selecionado normalmente.
        False -> a pesquisa não retornou nenhuma linha selecionável na
                 grade (`td.GridLink[onclick*='SelecionarItem']`); a
                 linha deve ser pulada e a automação segue para a próxima.
    """
    popup = abrir_popup(context, lambda: page.locator(seletor_botao_pesquisa).click())

    popup.locator("#btnCPF").click()
    popup.locator("#txtNuCpf").fill(cpf_formatado)

    mensagem_erro = popup.locator("td.SIGEFMensagemErro")
    resultado_grade = popup.locator("td.GridLink[onclick*='SelecionarItem']")

    popup.locator("#btnConfirmar").click()

    if not _aguardar_resultado_ou_erro_grade(resultado_grade, mensagem_erro):
        # Chega aqui quando a mensagem de erro apareceu (ou, no pior caso,
        # quando o timeout total expirou sem nenhum dos dois sinais) - a
        # checagem abaixo só deixa o aviso mais preciso quando o SIGEF
        # confirma explicitamente "Não há registros a serem listados.".
        if mensagem_erro.first.is_visible():
            print(
                f"⚠️  CPF {cpf_formatado}: SIGEF respondeu 'Não há registros a serem "
                f"listados.'. Pulando para a próxima linha."
            )
        else:
            print(f"⚠️  CPF {cpf_formatado}: nenhum registro encontrado no SIGEF. Pulando para a próxima linha.")
        popup.close()
        page.bring_to_front()
        return False

    resultado_grade.first.click()
    page.bring_to_front()
    return True


def _selecionar_empenho_por_numero(page: "Page", context: "BrowserContext", numero_empenho: str) -> bool:
    """
    Abre o popup de pesquisa de Nota de Empenho, pesquisa pelo número
    informado (já sem o prefixo "AAAANE") e seleciona o primeiro
    resultado da grade - mesmo padrão de `_selecionar_credor_por_cpf`
    (popup -> preencher -> confirmar -> selecionar item), usado pela
    automação NL.

    Retorna:
        True  -> empenho encontrado e selecionado normalmente.
        False -> a pesquisa não retornou nenhuma linha selecionável na
                 grade; a linha deve ser pulada e a automação segue para
                 a próxima.
    """
    popup = abrir_popup(context, lambda: page.locator("#txtNotaEmpenhoNumeroId_BtnPesquisa").click())

    mensagem_erro = popup.locator("td.SIGEFMensagemErro")
    resultado_grade = popup.locator("td.GridLink[onclick*='SelecionarItem']")

    popup.locator("#txtNotaEmpenhoNumero").fill(numero_empenho)
    popup.locator("#btnConfirmar").click()

    if not _aguardar_resultado_ou_erro_grade(resultado_grade, mensagem_erro):
        if mensagem_erro.first.is_visible():
            print(
                f"⚠️  NE {numero_empenho}: SIGEF respondeu 'Não há registros a serem "
                f"listados.'. Pulando para a próxima linha."
            )
        else:
            print(f"⚠️  NE {numero_empenho}: nenhum registro encontrado no SIGEF. Pulando para a próxima linha.")
        popup.close()
        page.bring_to_front()
        return False

    resultado_grade.first.click()
    page.bring_to_front()

    # O clique dispara um postback na tela principal (a NE selecionada volta
    # preenchida). Sem esperar esse retorno, a digitação do valor bruto
    # começava com a página ainda carregando e perdia dígitos.
    if not aguardar_campo_com_valor_por_id(page, "NotaEmpenho"):
        log_aviso(
            f"NE {numero_empenho}: a Nota de Empenho selecionada não voltou para a tela "
            f"principal. Pulando para a próxima linha."
        )
        return False

    aguardar_pagina_estavel(page)
    return True


def ce(dados, config=None, worksheet=None):
    """
    Automação: CE (Manter Despesa Certificada).

    Fluxo:
        1. Conecta 1 única vez à aba do SIGEF na tela de CE (parte fixa).
        2. Preenche os campos fixos da tela (gestão, mês de competência,
           observação, datas etc.) - também 1 única vez.
        3. Para cada linha da planilha: pesquisa o credor pelo CPF
           (coluna B), preenche o valor (coluna H, convertido para
           centavos), inclui o documento, captura o número da CE gerada
           e grava na coluna I da mesma linha.
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return

    if not dados:
        log_aviso("Nenhuma linha para processar na automação CE.")
        return

    config = config or {}
    linha_inicial = config.get("linha_inicial", 2)

    with sync_playwright() as p:
        try:
            context, page = conectar_e_obter_pagina_sigef(p, URL_CE_SIGEF, MARCADOR_URL_CE)
        except Exception as erro:
            log_erro(f"Erro ao conectar à tela de CE do SIGEF: {erro}")
            return

        try:
            _preencher_campos_fixos_ce(page, config)
        except Exception as erro:
            log_erro(f"Erro ao preencher os campos fixos da CE: {erro}")
            return

        # Locators reaproveitados em TODAS as linhas - buscados uma única
        # vez fora do loop, evitando repetir `page.locator(...)` a cada
        # iteração (reduz consultas ao DOM e chamadas ao Playwright).
        campo_valor = page.locator("#txtVlDocumento")
        botao_incluir = page.locator("#btnManutencao_BtnIncluir")
        campo_ce_gerada = page.locator("#txtNuSeq")

        total_processadas = 0

        # ---- Parte em loop até a última linha da planilha ------------------
        for indice, linha in enumerate(dados):
            numero_linha = linha_inicial + indice
            cpf = obter_celula(linha, colunas.COL_CE_CPF)
            valor_bruto = obter_celula(linha, colunas.COL_CE_VALOR)

            if not cpf or not valor_bruto:
                log_aviso(f"Linha {numero_linha}: CPF ou valor vazio. Pulando.")
                continue

            try:
                if not _selecionar_credor_por_cpf(page, context, formatar_cpf(cpf)):
                    continue

                # Requisito obrigatório: manter `press_sequentially()`
                # exatamente como está - não substituir por `.fill()`.
                campo_valor.press_sequentially(formatar_valor_centavos(valor_bruto))
                botao_incluir.click()

                # Captura a CE gerada e salva na coluna I da mesma linha.
                ce_gerada = campo_ce_gerada.input_value()
                log_sucesso(f"Linha {numero_linha}: CE '{ce_gerada}' gerada para o CPF {cpf}.")
                salvar_valor_gerado(worksheet, numero_linha, colunas.COL_CE_GERADA, ce_gerada, rotulo="CE")

                # Apaga a CE e o valor antes de reiniciar o loop na próxima linha.
                campo_ce_gerada.fill("")
                campo_valor.fill("")
                total_processadas += 1

            except Exception as erro:
                log_erro(f"Erro ao processar a linha {numero_linha} da automação CE: {erro}")
                continue

        log_sucesso(
            f"Automação CE concluída: {total_processadas} de {len(dados)} linha(s) processada(s)."
        )


def _extrair_ultima_despesa_certificada(page: "Page"):
    """
    Executa, dentro do próprio navegador (via `evaluate`, sem round-trips
    extras de Playwright para cada célula), a extração da última linha da
    grade `#dtgDespesaCertificada`: número da CE, mês/ano do documento e
    valor. Devolve os 3 valores já desempacotados (não o dicionário bruto
    retornado pelo JavaScript).
    """
    dados = page.locator("#dtgDespesaCertificada").evaluate("""
        table => {
            const linhas = [...table.querySelectorAll("tr")]
                .filter(tr => tr.querySelector("td.GridLink"));

            const ultima = linhas.at(-1);

            const numero = ultima.querySelector("td.GridLink").textContent.trim();

            const valor = [...ultima.querySelectorAll("td")]
                .find(td => td.align === "right")
                .textContent.trim();

            const numeroDocumento =
                ultima.querySelector("td.GridLink")
                    .nextElementSibling
                    .nextElementSibling
                    .textContent.trim();

            return { numero, numero_documento: numeroDocumento, valor };
        }
    """)
    return dados["numero"], dados["numero_documento"], dados["valor"]


def raspar_ce(dados, config=None, worksheet=None):
    """
    Automação: Raspar CE (Listar Despesa Certificada Geral).

    Mesma arquitetura da automação CE: conexão/preenchimento fixo 1 única
    vez, loop reaproveitando locators e reaproveitando o popup de busca de
    credor (`_selecionar_credor_por_cpf`, compartilhado com a CE).

    Fluxo:
        1. Conecta 1 única vez à aba do SIGEF na tela de listagem de CE
           (parte fixa).
        2. Preenche unidade gestora/gestão e abre a aba "Favorecida" -
           também 1 única vez.
        3. Para cada linha da planilha: pesquisa o credor pelo CPF
           (coluna B), confirma a busca, extrai da grade o número da CE,
           o mês/ano do documento e o valor da última despesa certificada
           listada, e grava os 3 valores na coluna I da mesma linha.
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return

    if not dados:
        log_aviso("Nenhuma linha para processar na automação Raspar CE.")
        return

    config = config or {}
    linha_inicial = config.get("linha_inicial", 2)

    with sync_playwright() as p:
        try:
            context, page = conectar_e_obter_pagina_sigef(p, URL_RASPAR_CE_SIGEF, MARCADOR_URL_RASPAR_CE)
        except Exception as erro:
            log_erro(f"Erro ao conectar à tela de listagem de CE do SIGEF: {erro}")
            return

        try:
            page.locator("#txtCdUnidadeGestora").fill("160001")
            page.locator("#txtCdGestao_SIGEFPesquisa").fill("00001")
            page.locator("#btnAbaFavorecida").click()
        except Exception as erro:
            log_erro(f"Erro ao preencher os campos fixos da Raspar CE: {erro}")
            return

        # Locators reaproveitados em TODAS as linhas - buscados uma única
        # vez fora do loop (reduz consultas ao DOM e chamadas ao Playwright).
        botao_confirmar = page.locator("#btnConfirmar")
        linhas_grade = page.locator("#dtgDespesaCertificada td.GridLink")

        total_processadas = 0

        # ---- Parte em loop até a última linha da planilha ------------------
        for indice, linha in enumerate(dados):
            numero_linha = linha_inicial + indice
            cpf = obter_celula(linha, colunas.COL_CE_CPF)

            if not cpf:
                log_aviso(f"Linha {numero_linha}: CPF vazio. Pulando.")
                continue

            try:
                # Mesmo popup de busca de credor da CE, só muda o botão
                # que o abre nesta tela (#txtCdCredor_BtnPesquisa).
                if not _selecionar_credor_por_cpf(
                    page, context, formatar_cpf(cpf), seletor_botao_pesquisa="#txtCdCredor_BtnPesquisa"
                ):
                    continue

                botao_confirmar.click()

                # Aguarda a grade recarregar (postback/AJAX) antes de ler os
                # dados - sem isso, a extração podia rodar antes da tabela
                # ser atualizada com o resultado da pesquisa.
                aguardar_pagina_estavel(page)
                linhas_grade.first.wait_for(state="visible", timeout=TIMEOUT_PADRAO_SIGEF)

                numero, numero_documento, valor = _extrair_ultima_despesa_certificada(page)
                print(f"Valores da linha [{numero_linha}]: {numero}, {numero_documento}, {valor}")

                # Grava só o número da CE na coluna I (mesma coluna/formato
                # que a automação CE grava) - a NL depende dessa coluna
                # conter apenas o número puro para poder pesquisar a
                # despesa certificada; mês/valor ficam só no log acima.
                salvar_valor_gerado(worksheet, numero_linha, colunas.COL_CE_GERADA, numero, rotulo="Raspar CE")

                total_processadas += 1

            except Exception as erro:
                log_erro(f"Erro ao processar a linha {numero_linha} da automação Raspar CE: {erro}")
                continue

        log_sucesso(
            f"Automação Raspar CE concluída: {total_processadas} de {len(dados)} linha(s) processada(s)."
        )


