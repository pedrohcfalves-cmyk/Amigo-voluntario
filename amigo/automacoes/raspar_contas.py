"""Automação Raspar contas (confere banco/agência/conta gravados no SIGEF)."""
from .. import colunas
from ..constantes import URL_RASPAR_CONTA_SIGEF, MARCADOR_URL_RASPAR_CONTA
from ..excel import obter_celula, atualizar_status, salvar_valor_gerado
from ..log import log_info, log_sucesso, log_erro, log_aviso
from ..navegador import conectar_e_obter_pagina_sigef, aguardar_pagina_estavel, fechar_paginas, abrir_popup
from ..playwright_compat import sync_playwright, PlaywrightTimeoutError
from ..utils import extrair_numero_documento, normalizar_chave_bancaria, formatar_valor_br
from .pp import ler_tabela_pp_geral, _valor_planilha_para_float, _linha_grade_tem_valor

def raspar_conta(dados, config=None, worksheet=None):
    """
    Automação: Raspar contas (Listar Preparação Pagamento Geral).

    Confere se o domicílio bancário que ficou gravado em cada PP no SIGEF
    é o mesmo que está na planilha - ou seja, é a conferência DEPOIS do
    fato, complementando a validação que a automação PP faz ANTES de
    gerar o documento.

    Fluxo:
        1. Conecta 1 única vez à tela "Listar Preparação Pagamento Geral".
        2. Para cada linha da planilha:
           a. Preenche a parte fixa (unidade gestora 160001 / gestão 00001)
              e o número da PP da coluna K, já sem o prefixo "AAAAPP", e
              confirma.
           b. Lê a grade `#dtgPP` e procura a linha cujo valor bate com a
              coluna H. Se NENHUMA bater, usa a PRIMEIRA linha da grade -
              a pesquisa já foi filtrada pelo número exato da PP, então a
              grade normalmente traz só ela; o casamento pelo valor é uma
              conferência a mais, não um filtro que possa descartar a
              linha certa.
           c. Abre o detalhe da PP e captura banco, agência e conta.
           d. Grava os três na coluna M e escreve "Igual"/"Diferente" na
              coluna N, comparando com as colunas E, F e G.
        3. Erro em uma linha nunca interrompe a automação - registra e
           segue para a próxima.

    A comparação usa `normalizar_chave_bancaria` nos dois lados (nunca
    texto bruto): a planilha traz "1 / 0951-2 / 81.590-X" e o SIGEF
    devolve "001 / 009512 / 000081590X".
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return

    if not dados:
        log_aviso("Nenhuma linha para processar na automação Raspar contas.")
        return

    config = config or {}
    linha_inicial = config.get("linha_inicial", 2)

    with sync_playwright() as p:
        try:
            context, page = conectar_e_obter_pagina_sigef(
                p, URL_RASPAR_CONTA_SIGEF, MARCADOR_URL_RASPAR_CONTA
            )
        except Exception as erro:
            log_erro(f"Erro ao conectar à tela de listagem de PP do SIGEF: {erro}")
            return

        # Locators reaproveitados em TODAS as linhas - buscados uma única
        # vez fora do loop (reduz consultas ao DOM).
        campo_unidade_gestora = page.locator("#txtUnidadeGestora")
        campo_gestao = page.locator("#txtGestao_SIGEFPesquisa")
        campo_pp = page.locator("#txtPrepPagSeq")
        botao_confirmar = page.locator("#btnConfirmar")
        linhas_grid_link = page.locator("#dtgPP td.GridLink")

        total_processadas = 0
        total_iguais = 0
        total_diferentes = 0

        for indice, linha in enumerate(dados):
            numero_linha = linha_inicial + indice

            pp_bruta = obter_celula(linha, colunas.COL_RASPAR_CONTA_PP)
            if not pp_bruta:
                log_aviso(f"Linha {numero_linha}: PP (coluna K) vazia. Pulando.")
                continue

            pp_pesquisa = extrair_numero_documento(pp_bruta).zfill(6)

            # O valor da coluna H é opcional aqui: serve só para escolher a
            # linha certa quando a grade traz mais de uma. Sem ele (ou se
            # não for numérico), cai no comportamento padrão de usar a
            # primeira linha.
            valor_bruto = obter_celula(linha, colunas.COL_RASPAR_CONTA_VALOR)
            try:
                valor_esperado = _valor_planilha_para_float(valor_bruto) if valor_bruto else None
            except (ValueError, TypeError):
                valor_esperado = None

            log_info(f"Linha {numero_linha}: conferindo a PP {pp_bruta}...")

            try:
                # Parte fixa preenchida A CADA linha (e não uma única vez
                # fora do loop): a tela faz postback ao confirmar, e os
                # campos podem voltar em branco. Preencher de novo é barato
                # e evita pesquisar na unidade gestora errada.
                campo_unidade_gestora.fill("160001")
                campo_gestao.fill("00001")
                campo_pp.fill(pp_pesquisa)

                botao_confirmar.click()
                aguardar_pagina_estavel(page)

                registros = ler_tabela_pp_geral(page)
                if not registros:
                    log_aviso(
                        f"Linha {numero_linha}: grade vazia para a PP "
                        f"{pp_bruta}. Pulando para a próxima linha."
                    )
                    continue

                escolhido = None
                if valor_esperado is not None:
                    escolhido = next(
                        (
                            registro for registro in registros
                            if len(registro["celulas"]) > 1
                            and _linha_grade_tem_valor(registro["celulas"], valor_esperado)
                        ),
                        None,
                    )

                if escolhido is None:
                    # Fallback previsto: usa a primeira linha da grade.
                    escolhido = registros[0]
                    if valor_esperado is not None:
                        log_aviso(
                            f"Linha {numero_linha}: nenhuma linha da grade com o valor "
                            f"{formatar_valor_br(valor_bruto)}. Usando a primeira linha "
                            f"({escolhido['texto']})."
                        )

                detalhe = abrir_popup(context, lambda: linhas_grid_link.nth(escolhido["indice"]).click())

                banco = detalhe.locator("#txtDomicilioBancarioBanco").input_value()
                agencia = detalhe.locator("#txtDomicilioBancarioAgencia").input_value()
                conta = detalhe.locator("#txtDomicilioBancarioConta").input_value()

                try:
                    detalhe.locator("#SIGEFBotoesImpressao_BtnFechar").click()
                except Exception:
                    # O botão pode não estar presente/visível; o
                    # `fechar_paginas` abaixo encerra a aba de qualquer forma.
                    pass
                fechar_paginas([detalhe])
                page.bring_to_front()

                # ---- Grava o que o SIGEF devolveu (coluna M) --------------
                domicilio_sigef = f"{banco} / {agencia} / {conta}"
                salvar_valor_gerado(
                    worksheet, numero_linha, colunas.COL_RASPAR_CONTA_RESULTADO,
                    domicilio_sigef, rotulo="Domicílio (SIGEF)"
                )

                # ---- Compara com a planilha (colunas E, F e G) ------------
                banco_planilha = obter_celula(linha, colunas.COL_RASPAR_CONTA_BANCO)
                agencia_planilha = obter_celula(linha, colunas.COL_RASPAR_CONTA_AGENCIA)
                conta_planilha = obter_celula(linha, colunas.COL_RASPAR_CONTA_CONTA)

                divergencias = [
                    rotulo
                    for rotulo, do_sigef, da_planilha in (
                        ("banco", banco, banco_planilha),
                        ("agência", agencia, agencia_planilha),
                        ("conta", conta, conta_planilha),
                    )
                    if normalizar_chave_bancaria(do_sigef)
                    != normalizar_chave_bancaria(da_planilha)
                ]

                if divergencias:
                    resultado = "Diferente"
                    total_diferentes += 1
                    log_aviso(
                        f"Linha {numero_linha}: DIFERENTE em {', '.join(divergencias)}. "
                        f"SIGEF: {domicilio_sigef} | planilha: "
                        f"{banco_planilha} / {agencia_planilha} / {conta_planilha}"
                    )
                else:
                    resultado = "Igual"
                    total_iguais += 1
                    log_sucesso(f"Linha {numero_linha}: IGUAL ({domicilio_sigef}).")

                atualizar_status(
                    worksheet, numero_linha, colunas.COL_RASPAR_CONTA_CONFERENCIA, resultado
                )

                total_processadas += 1

            except PlaywrightTimeoutError as erro:
                log_erro(f"Timeout ao processar a linha {numero_linha} da Raspar contas: {erro}")
                continue
            except Exception as erro:
                log_erro(f"Erro ao processar a linha {numero_linha} da Raspar contas: {erro}")
                continue

        log_sucesso(
            f"Automação Raspar contas concluída: {total_processadas} de "
            f"{len(dados)} linha(s) processada(s) - {total_iguais} igual(is) e "
            f"{total_diferentes} diferente(s)."
        )


