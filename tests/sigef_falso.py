# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
"SIGEF DE MENTIRA" PARA OS TESTES
=================================
Imita o pedaço do Playwright que as automações usam (navegador, abas,
campos, botões, popups), respondendo como o SIGEF responderia - sem
navegador e sem internet. Assim os testes rodam o código REAL das
automações CE, NL, PP e OB e conferem, por exemplo, que no modo simulação
o clique final de confirmar NUNCA acontece.

Tudo que a automação faz fica registrado em `SigefFalso.acoes`
(("click", seletor), ("fill", seletor, valor), ("goto", url)...).
"""
from contextlib import contextmanager

DOMINIO = "sigef.sefin.ro.gov.br"


class SigefFalso:
    """Estado compartilhado de todas as abas/popups de 1 teste."""

    def __init__(self, ano="2027", credor_existe=True):
        self.ano = ano
        self.credor_existe = credor_existe
        self.acoes = []
        self.urls_abertas = []

    def cliques(self):
        return [acao[1] for acao in self.acoes if acao[0] == "click"]

    # Respostas das telas (o que o JavaScript das automações "leria").
    def documentos_nl(self, pagina):
        if pagina.confirmado:
            return [{"documento": f"{self.ano}NL000222", "valor_liquido": "1.400,50"}]
        return []

    def ordem_cronologica(self):
        return [{"indice": 0, "texto": f"{self.ano}NL000222 {self.ano}CE000111 1.400,50"}]

    def domicilios(self):
        return [{"indice": 0, "banco": "001", "agencia": "01178-9", "conta": "0000745081"}]

    def grade_ob(self):
        return [{"indice": 0, "pp": f"{self.ano}PP000333", "cpf": "529.982.247-25", "valor": "1.400,50"}]


class LocalizadorFalso:
    def __init__(self, pagina, seletor):
        self.pagina = pagina
        self.seletor = seletor

    @property
    def first(self):
        return self

    def nth(self, _indice):
        return self

    def locator(self, sub):
        return LocalizadorFalso(self.pagina, f"{self.seletor} {sub}")

    # ---- ações ----
    def fill(self, valor, **_kw):
        self.pagina.sigef.acoes.append(("fill", self.seletor, valor))
        self.pagina.valores[self.seletor] = valor

    def press_sequentially(self, valor, **_kw):
        self.pagina.sigef.acoes.append(("digitar", self.seletor, valor))
        self.pagina.valores[self.seletor] = valor

    def click(self, **_kw):
        self.pagina.sigef.acoes.append(("click", self.seletor))
        self.pagina.ao_clicar(self.seletor)

    def check(self, **_kw):
        self.pagina.sigef.acoes.append(("check", self.seletor))
        self.pagina.marcados[self.seletor] = True

    def select_option(self, value=None, *_a, **_kw):
        self.pagina.sigef.acoes.append(("select", self.seletor, value))
        self.pagina.valores[self.seletor] = value

    def wait_for(self, **_kw):
        return None

    # ---- leituras ----
    def input_value(self, **_kw):
        if self.seletor in self.pagina.valores:
            return self.pagina.valores[self.seletor]
        return self.pagina.valor_inicial(self.seletor)

    def count(self):
        if "MensagemErro" in self.seletor:
            return 0
        if "GridLink" in self.seletor and self.pagina.eh_popup_credor and not self.pagina.sigef.credor_existe:
            return 0
        return 1

    def is_visible(self):
        # "Não há registros a serem listados." na pesquisa de credor.
        return ("MensagemErro" in self.seletor and self.pagina.eh_popup_credor
                and not self.pagina.sigef.credor_existe)

    def is_checked(self):
        return self.pagina.marcados.get(self.seletor, False)

    def inner_text(self):
        if "MensagemSucesso" in self.seletor:
            return self.pagina.mensagem_sucesso
        return ""

    def evaluate(self, script, *_args):
        # O JavaScript das automações não traz o nome da tabela - ele vem do
        # seletor do locator (ex: "#dtgDocumentos").
        return self.pagina.avaliar(f"{self.seletor} {script}")


class PaginaFalsa:
    def __init__(self, sigef, url="", eh_popup_credor=False):
        self.sigef = sigef
        self.url = url
        self.valores = {}
        self.marcados = {}
        self.confirmado = False
        self.mensagem_sucesso = ""
        self.eh_popup_credor = eh_popup_credor
        self.fechada = False

    def locator(self, seletor):
        return LocalizadorFalso(self, seletor)

    def valor_inicial(self, seletor):
        if seletor == "#txtNuSeq":  # número da CE depois do "Incluir"
            return f"{self.sigef.ano}CE000111" if ("click", "#btnManutencao_BtnIncluir") in self.sigef.acoes else ""
        return ""

    def ao_clicar(self, seletor):
        if seletor == "#btnConfirmar":
            self.confirmado = True
        if seletor.startswith("#chk"):
            self.marcados[seletor] = True
        if seletor == "#btnConfirmar" or seletor == "#btnManutencao_BtnIncluir":
            self.mensagem_sucesso = f"Operação realizada. O número gerado foi {self.sigef.ano}PP000333."
        if seletor == "#SIGEFBotoesManutencao_BtnIncluir":
            self.mensagem_sucesso = f"Operação realizada. O número gerado foi {self.sigef.ano}OB000444."

    def avaliar(self, script):
        if "dtgDocumentos" in script:
            return self.sigef.documentos_nl(self)
        if "readyState" in script:
            return True
        if "input[type=text]" in script:
            return True
        if "divdtgGerarOrdemCronologica" in script:
            return self.sigef.ordem_cronologica()
        if "dtgDomicilioBancario" in script:
            return self.sigef.domicilios()
        if "GridLinhaPar" in script:
            return self.sigef.grade_ob()
        if "dtgDespesaCertificada" in script or "dtgPP" in script:
            return []
        return None

    def evaluate(self, script, *_args):
        return self.avaliar(script)

    # ---- navegação/espera ----
    def goto(self, url, **_kw):
        self.sigef.acoes.append(("goto", url))
        self.sigef.urls_abertas.append(url)
        self.url = url

    def wait_for_load_state(self, *_a, **_kw):
        return None

    def wait_for_selector(self, *_a, **_kw):
        return None

    def wait_for_event(self, *_a, **_kw):
        return None

    def wait_for_timeout(self, *_a, **_kw):
        return None

    def bring_to_front(self):
        return None

    def is_closed(self):
        return self.fechada

    def close(self):
        self.fechada = True


class _InfoPopup:
    def __init__(self, pagina):
        self.value = pagina


class ContextoFalso:
    def __init__(self, sigef):
        self.sigef = sigef
        # Uma aba do SIGEF já aberta no PORTAL do ano anterior: a automação
        # precisa navegar para a tela certa, no exercício certo.
        self.pages = [PaginaFalsa(sigef, url=f"http://{DOMINIO}/SIGEF2026/SIGEFPortal.html")]

    def new_page(self):
        pagina = PaginaFalsa(self.sigef)
        self.pages.append(pagina)
        return pagina

    @contextmanager
    def expect_page(self, **_kw):
        # O popup que abrir agora é o de pesquisa de credor? (para simular
        # "CPF não encontrado").
        ultimo_clique = self.sigef.cliques()[-1] if self.sigef.cliques() else ""
        popup = PaginaFalsa(self.sigef, url=f"http://{DOMINIO}/popup",
                            eh_popup_credor=False)
        info = _InfoPopup(popup)
        yield info
        clique = self.sigef.cliques()[-1] if self.sigef.cliques() else ultimo_clique
        popup.eh_popup_credor = "Credor" in clique


class NavegadorFalso:
    def __init__(self, sigef):
        self.contexts = [ContextoFalso(sigef)]


class _Chromium:
    def __init__(self, sigef):
        self.sigef = sigef

    def connect_over_cdp(self, _endereco):
        return NavegadorFalso(self.sigef)


class _PlaywrightFalso:
    def __init__(self, sigef):
        self.chromium = _Chromium(sigef)


def sync_playwright_falso(sigef):
    """Substituto de `playwright.sync_api.sync_playwright` para 1 teste."""
    @contextmanager
    def _fabrica():
        yield _PlaywrightFalso(sigef)
    return _fabrica
