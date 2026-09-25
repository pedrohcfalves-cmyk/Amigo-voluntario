"""
PLAYWRIGHT - UTILIDADES GENÉRICAS (conexão via CDP)
=====================================================
Funções reaproveitadas por qualquer automação do SIGEF: abrir/conectar o
navegador, esperar a página ficar estável, clicar e aguardar navegação,
preencher e conferir campos, etc. Independentes da automação específica.
"""
import time

from .constantes import DOMINIO_SIGEF, TIMEOUT_PADRAO_SIGEF, TIMEOUT_BOTAO_LIMPAR, REGEX_APENAS_DIGITOS
from .log import log_info, log_sucesso, log_erro, log_aviso
from .playwright_compat import sync_playwright, Page, BrowserContext, PlaywrightTimeoutError

def iniciar_navegador():
    """
    Inicia um navegador novo via Playwright (não depende de um Chrome já
    aberto). Retorna (playwright, browser, page) para que o chamador possa
    controlar o ciclo de vida (e fechar corretamente ao final).

    Uso típico:
        pw, browser, page = iniciar_navegador()
        try:
            page.goto("https://sigef.exemplo.gov.br")
            ...
        finally:
            fechar_navegador(pw, browser)
    """
    if sync_playwright is None:
        log_erro("A biblioteca 'playwright' não está instalada.")
        log_info("Instale com: pip install playwright && playwright install")
        return None, None, None

    try:
        pw = sync_playwright().start()
        browser = pw.chromium.launch(headless=False)
        page = browser.new_page()
        log_sucesso("Navegador iniciado com sucesso.")
        return pw, browser, page
    except Exception as erro:
        log_erro(f"Erro ao iniciar o navegador: {erro}")
        return None, None, None


def fechar_navegador(pw, browser):
    """Fecha o navegador e encerra a instância do Playwright com segurança."""
    try:
        if browser is not None:
            browser.close()
        if pw is not None:
            pw.stop()
        log_info("Navegador encerrado.")
    except Exception as erro:
        log_erro(f"Erro ao fechar o navegador: {erro}")


def conectar_navegador_sigef(playwright):
    """
    Conecta ao Chrome já aberto em modo de depuração (porta 9222). Todas as
    automações do SIGEF usam esta MESMA conexão, pois dependem do usuário
    já estar autenticado no SIGEF nesse Chrome. Deve ser chamada 1 única
    vez por execução.
    """
    try:
        browser = playwright.chromium.connect_over_cdp("http://127.0.0.1:9222")
    except Exception:
        raise Exception(
            "Chrome em modo de depuração não encontrado.\n"
            "Abra o Chrome com:\n"
            "\"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe\" "
            "--remote-debugging-port=9222"
        )
    return browser


def localizar_ou_abrir_pagina_sigef(context: "BrowserContext", url: str, url_marker: str) -> "Page":
    """
    Implementação genérica, reaproveitada por TODAS as automações do
    SIGEF: reaproveita a aba já aberta no domínio do SIGEF (evita abrir
    uma nova aba/contexto a cada automação - reduz chamadas ao
    navegador), navegando para `url` somente se a aba encontrada não for
    a tela esperada (`url_marker` ausente da URL atual). Se nenhuma aba
    do SIGEF estiver aberta, cria uma nova.

    Cada automação mantém sua própria função `localizar_ou_abrir_pagina_*`
    (fina, delegando para esta) para preservar nomes descritivos nos
    pontos de chamada e permitir evolução independente, se necessário.

    Comportamento (qualquer aba do domínio SIGEF é "tomada"):
        - Se já existir uma aba EXATAMENTE na tela esperada (`url_marker`
          bate), reaproveita ela direto, sem navegar de novo.
        - Caso contrário, toma controle de QUALQUER outra aba do domínio
          SIGEF que esteja aberta (ex: o portal `SIGEFPortal.html`, ou a
          tela de outra automação que rodou antes) e a substitui pela URL
          pedida agora - e vice-versa: ao trocar de automação, a aba que
          antes estava na tela da CE, por exemplo, é redirecionada para a
          tela da nova automação.
        - Se não houver NENHUMA aba do SIGEF aberta, cria uma nova.

    Também fecha, ao final, qualquer OUTRA aba/popup do domínio do SIGEF
    que não seja a aba escolhida acima - evita que a automação confunda a
    aba principal com um popup órfão de uma execução anterior (ex:
    "Pesquisar Credor" que ficou aberto), o que travava o
    `context.expect_page()` das automações esperando um evento de "nova
    página" que nunca chegava.
    """
    page = None
    candidata_generica = None

    for aba in context.pages:
        if DOMINIO_SIGEF not in aba.url.lower():
            continue
        if url_marker in aba.url:
            page = aba
            break
        if candidata_generica is None:
            candidata_generica = aba

    if page is None:
        page = candidata_generica

    if page is None:
        page = context.new_page()
        page.goto(url)
    else:
        page.bring_to_front()
        if url_marker not in page.url:
            page.goto(url)

    for aba in list(context.pages):
        if aba is not page and DOMINIO_SIGEF in aba.url.lower():
            try:
                aba.close()
            except Exception:
                pass

    # `aguardar_pagina_estavel` no lugar do `networkidle` puro: sem teto, ele
    # trava até o timeout em telas do SIGEF que mantêm requisições em segundo
    # plano (otimização trazida da automação PP).
    aguardar_pagina_estavel(page)
    return page


def conectar_e_obter_pagina_sigef(playwright, url: str, url_marker: str):
    """
    Conecta ao Chrome em modo debug, garante que existe ao menos um
    BrowserContext e localiza/abre a aba do SIGEF na URL informada.
    Concentra em 1 único lugar a sequência repetida em toda automação:
    conectar -> validar contexto -> obter página.

    Chamada dentro de um `with sync_playwright() as p:` já aberto pela
    automação (o navegador/contexto permanecem sob controle dela).
    """
    browser = conectar_navegador_sigef(playwright)
    if not browser.contexts:
        raise Exception("Nenhum contexto de navegador encontrado no Chrome conectado.")
    context = browser.contexts[0]
    page = localizar_ou_abrir_pagina_sigef(context, url, url_marker)
    return context, page


def aguardar_pagina_estavel(page: "Page", timeout_networkidle: int = 5000) -> None:
    """
    Aguarda o DOM carregar (`domcontentloaded`, bloqueante) e, na
    sequência, tenta aguardar a rede ficar ociosa (`networkidle`) por até
    `timeout_networkidle` ms - sem falhar caso a página tenha pollers/
    requisições contínuas que nunca ficam totalmente ociosas (comum em
    telas do SIGEF).
    """
    page.wait_for_load_state("domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=timeout_networkidle)
    except Exception:
        pass


def abrir_popup(context: "BrowserContext", acao_de_clique) -> "Page":
    """
    Executa `acao_de_clique()` (um clique que abre uma página/popup nova do
    SIGEF) dentro de `context.expect_page()`, aguarda a nova página
    estabilizar (`aguardar_pagina_estavel`) e a devolve.

    Reaproveitado por todas as automações que abrem um popup e já esperam
    a rede ficar ociosa antes do próximo passo - padrão idêntico repetido
    em CE, NL, PP, Raspar contas e Gerar OB. Automações que, em vez disso,
    esperam por um ELEMENTO específico do popup (mais rápido e mais
    confiável em telas mais lentas) continuam abrindo o popup manualmente
    com `context.expect_page()`, sem usar este helper.
    """
    with context.expect_page() as popup_info:
        acao_de_clique()
    popup = popup_info.value
    aguardar_pagina_estavel(popup)
    return popup


def aguardar_pagina_ociosa(page: "Page", periodo_calmo: float = 0.6,
                           timeout: int = TIMEOUT_PADRAO_SIGEF) -> bool:
    """
    Espera a página parar DE VERDADE: `document.readyState == "complete"`
    mantido por `periodo_calmo` segundos seguidos.

    `aguardar_pagina_estavel()` sozinho não basta quando o SIGEF encadeia
    postbacks: o DOM fica "pronto" por um instante e volta a recarregar
    logo em seguida. Uma ação disparada nessa janela (selecionar o combo,
    por exemplo) é desfeita pelo recarregamento seguinte.

    Retorna True se a página ficou parada dentro do prazo.
    """
    aguardar_pagina_estavel(page)

    prazo = time.monotonic() + (timeout / 1000)
    calmo_desde = None

    while time.monotonic() < prazo:
        try:
            pronta = page.evaluate("() => document.readyState === 'complete'")
        except Exception:
            # navegação em curso: o contexto de execução foi destruído
            pronta = False
            calmo_desde = None

        if pronta:
            if calmo_desde is None:
                calmo_desde = time.monotonic()
            elif time.monotonic() - calmo_desde >= periodo_calmo:
                return True
        else:
            calmo_desde = None

        time.sleep(0.1)

    return False


def clicar_e_aguardar_elemento(page: "Page", seletor: str, seletor_esperado: str = None,
                               timeout: int = TIMEOUT_PADRAO_SIGEF) -> None:
    """
    Clica em `seletor` e, em vez de esperar a rede ficar ociosa
    (`networkidle` - que em telas do SIGEF com requisições em segundo
    plano fica pendurado até estourar o timeout), espera apenas o
    ELEMENTO SEGUINTE do fluxo aparecer: é ele que prova que o postback
    terminou.

    Passe `seletor_esperado=None` quando o passo seguinte já tiver a sua
    própria espera (ex: o `wait_for_selector` da mensagem de retorno do
    SIGEF) - assim nenhuma espera é feita em duplicidade.
    """
    alvo = page.locator(seletor).first
    alvo.wait_for(state="visible", timeout=timeout)
    alvo.click(timeout=timeout)

    if seletor_esperado:
        page.locator(seletor_esperado).first.wait_for(state="visible", timeout=timeout)


def limpar_formulario(page: "Page", timeout: int = TIMEOUT_BOTAO_LIMPAR) -> bool:
    """
    Devolve o formulário ao estado inicial clicando em "Limpar", com um
    plano B para o caso de o botão não estar na tela.

    Em algumas telas do SIGEF (e depois de certos erros), a automação
    fica numa etapa em que o "Limpar" simplesmente não existe - só o
    "Voltar". Antes, isso estourava um timeout e a linha seguinte
    começava com a tela suja. Agora, SE E SOMENTE SE der timeout
    esperando o "Limpar", clica em `#btnVoltar` e tenta o "Limpar"
    novamente.

    Compartilhada por todas as automações que limpam o formulário entre
    uma linha e outra (NL, PP, ...).

    Retorna True se conseguiu limpar; False se nem pelo Voltar deu.
    """
    seletor_limpar = "img[src*='Limpar.GIF']"

    try:
        botao_limpar = page.locator(seletor_limpar).first
        botao_limpar.wait_for(state="visible", timeout=timeout)
        botao_limpar.click(timeout=timeout)
        return True
    except PlaywrightTimeoutError:
        log_aviso("Botão 'Limpar' não encontrado na tela. Tentando pelo 'Voltar'.")
    except Exception as erro:
        log_aviso(f"Não foi possível clicar em 'Limpar': {erro}")
        return False

    try:
        botao_voltar = page.locator("#btnVoltar").first
        botao_voltar.wait_for(state="visible", timeout=timeout)
        botao_voltar.click(timeout=timeout)

        botao_limpar = page.locator(seletor_limpar).first
        botao_limpar.wait_for(state="visible", timeout=timeout)
        botao_limpar.click(timeout=timeout)
        log_info("Formulário limpo após voltar para a tela anterior.")
        return True
    except Exception as erro:
        log_erro(f"Não foi possível limpar o formulário nem pelo 'Voltar': {erro}")
        return False


def aguardar_campo_com_valor_por_id(page: "Page", padrao_id: str,
                                    timeout: int = TIMEOUT_PADRAO_SIGEF) -> bool:
    """
    Espera até que exista, na tela, um campo de texto cujo `id` case com
    `padrao_id` (sem diferenciar maiúsculas) e que esteja PREENCHIDO.

    É o sinal de que um popup de pesquisa do SIGEF terminou de devolver o
    registro escolhido para a tela principal - o clique na grade dispara
    um postback, e a tela principal só fica utilizável quando ele acaba.
    A checagem é feita por JavaScript, sem depender do `id` exato (o
    SIGEF varia o sufixo do campo entre telas).

    Exemplos de uso: "NotaLancamento" (PP) e "NotaEmpenho" (NL).
    """
    prazo = time.monotonic() + (timeout / 1000)

    while time.monotonic() < prazo:
        try:
            preenchido = page.evaluate("""
                padrao => [...document.querySelectorAll("input[type=text]")]
                    .some(campo => new RegExp(padrao, "i").test(campo.id)
                                   && campo.value.trim() !== "")
            """, padrao_id)
            if preenchido:
                return True
        except Exception:
            pass
        time.sleep(0.1)

    return False


def aguardar_campo_preenchido(campo, valor_esperado: str,
                              timeout: int = TIMEOUT_PADRAO_SIGEF) -> bool:
    """
    Confere se um campo terminou de receber TODO o valor digitado, antes
    de a automação seguir em frente.

    Necessário para os campos preenchidos com `press_sequentially()`: a
    digitação é tecla a tecla e as máscaras/JS do SIGEF podem atrasar (ou
    engolir) os últimos dígitos - o clique seguinte então acontecia com o
    campo pela metade.

    A comparação é feita só pelos dígitos, já que a máscara do SIGEF
    reformata o que foi digitado (ex: digitado "140000", exibido
    "1.400,00" - os dois viram "140000").
    """
    digitos_esperados = REGEX_APENAS_DIGITOS.sub("", str(valor_esperado)).lstrip("0")
    prazo = time.monotonic() + (timeout / 1000)

    while time.monotonic() < prazo:
        try:
            digitos_na_tela = REGEX_APENAS_DIGITOS.sub("", campo.input_value()).lstrip("0")
            if digitos_na_tela == digitos_esperados:
                return True
        except Exception:
            pass
        time.sleep(0.1)

    return False


def selecionar_combo_com_conferencia(page: "Page", combo, valor: str, tentativas: int = 3,
                                     timeout: int = TIMEOUT_PADRAO_SIGEF) -> bool:
    """
    Seleciona uma opção de <select> pelo VALUE e CONFERE se ela ficou
    mesmo selecionada DEPOIS de a página parar de recarregar, repetindo
    enquanto necessário.

    Combos do SIGEF são montados e revalidados por postback. Havia duas
    formas de perder a seleção:
      - selecionar antes de a lista terminar de montar; e
      - selecionar no meio de um postback ainda em curso, que logo em
        seguida recarrega a tela e desfaz a escolha (era o caso do Tipo
        de Ordem Bancária, que "sumia" e a automação seguia para a
        pesquisa da conta sem tipo nenhum).

    Por isso a página é aguardada parada (`aguardar_pagina_ociosa`) antes
    de selecionar E depois de selecionar - só então o valor é lido de
    volta. Só retorna True com o combo de fato no valor pedido.
    """
    for _ in range(tentativas):
        try:
            aguardar_pagina_ociosa(page, timeout=timeout)

            combo.wait_for(state="visible", timeout=timeout)
            combo.locator(f"option[value='{valor}']").first.wait_for(state="attached", timeout=timeout)
            combo.select_option(value=valor)

            # A própria seleção pode disparar postback: espera a tela parar
            # de novo antes de conferir.
            aguardar_pagina_ociosa(page, timeout=timeout)

            if combo.input_value() == valor:
                return True
        except Exception:
            pass

        time.sleep(0.2)

    return False


def fechar_paginas(paginas) -> None:
    """
    Fecha somente as janelas/popups que a automação abriu (ignora as
    demais abas do usuário e as que o próprio SIGEF já fechou). Evita que
    um popup órfão trave o `context.expect_page()` da linha seguinte.
    """
    for pagina in paginas:
        try:
            if pagina is not None and not pagina.is_closed():
                pagina.close()
        except Exception:
            pass



