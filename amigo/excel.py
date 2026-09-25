"""
EXCEL (pywin32/COM)
====================
Conecta no Microsoft Excel já aberto na tela do usuário (ou abre, se
preciso), lê os dados da planilha e grava os resultados das automações -
tudo de forma VISÍVEL, para acompanhamento em tempo real.
"""
import os
import time
from typing import List, Optional

try:
    import win32com.client as win32
except ImportError:
    win32 = None

from . import colunas
from .log import log_info, log_sucesso, log_erro, log_aviso

XL_UP = -4162        # constante COM: Excel.Constants.xlUp
XL_TO_LEFT = -4159   # constante COM: Excel.Constants.xlToLeft


def _obter_excel_em_execucao():
    """
    Tenta obter uma instância do Excel que já esteja aberta na tela do
    usuário (via GetActiveObject). Retorna None se nenhuma instância do
    Excel estiver rodando no momento.
    """
    if win32 is None:
        return None
    try:
        return win32.GetActiveObject("Excel.Application")
    except Exception:
        return None


def _localizar_pasta_de_trabalho_aberta(excel_app, caminho_planilha: str):
    """
    Procura, entre as pastas de trabalho já abertas na instância do Excel
    informada, uma cujo caminho completo corresponda a `caminho_planilha` -
    para reaproveitar o arquivo já aberto em vez de abri-lo de novo.
    """
    if excel_app is None or not caminho_planilha:
        return None

    caminho_normalizado = os.path.normcase(os.path.abspath(caminho_planilha))
    try:
        for pasta in excel_app.Workbooks:
            try:
                if os.path.normcase(os.path.abspath(pasta.FullName)) == caminho_normalizado:
                    return pasta
            except Exception:
                continue
    except Exception:
        pass
    return None


def conectar_planilha(config: dict):
    """
    Conecta ao arquivo Excel configurado, via COM (pywin32).

    Comportamento (equivalente ao antigo fluxo do Google Sheets, porém
    tudo em uma única etapa, já que aqui não existe distinção entre "abrir
    para visualizar" e "conectar para ler/escrever" - é o mesmo Excel):
        - Se o arquivo já estiver aberto em uma instância do Excel na tela
          do usuário, essa instância/aba é REAPROVEITADA (não abre de novo).
        - Caso contrário, abre uma nova instância do Excel, VISÍVEL, com o
          arquivo configurado.

    Requisitos:
        - Microsoft Excel instalado (Windows).
        - 'caminho_planilha' definido nas configurações.

    Retorna:
        worksheet (objeto COM Worksheet) em caso de sucesso, ou None em
        caso de erro.
    """
    if win32 is None:
        log_erro("A biblioteca 'pywin32' não está instalada.")
        log_info("Instale com: pip install pywin32")
        return None

    caminho_planilha = config.get("caminho_planilha")
    if not caminho_planilha:
        log_erro("Nenhum arquivo Excel configurado. Use a opção 'Alterar planilha' no menu.")
        return None

    if not os.path.exists(caminho_planilha):
        log_erro(f"Arquivo não encontrado: {caminho_planilha}")
        return None

    try:
        excel_app = _obter_excel_em_execucao()
        pasta = _localizar_pasta_de_trabalho_aberta(excel_app, caminho_planilha) if excel_app else None

        if pasta is None:
            if excel_app is None:
                excel_app = win32.Dispatch("Excel.Application")
            excel_app.Visible = True
            pasta = excel_app.Workbooks.Open(os.path.abspath(caminho_planilha))
            log_sucesso(f"Arquivo Excel aberto: {os.path.basename(caminho_planilha)}")
        else:
            excel_app.Visible = True
            log_sucesso(f"Reaproveitando arquivo Excel já aberto: {os.path.basename(caminho_planilha)}")

        aba_nome = config.get("aba_planilha")
        if aba_nome:
            worksheet = pasta.Worksheets(aba_nome)
        else:
            worksheet = pasta.Worksheets(1)

        worksheet.Activate()
        excel_app.WindowState = -4137  # xlMaximized: traz o Excel para frente, maximizado
        log_sucesso(f"Conectado à aba '{worksheet.Name}'.")
        return worksheet

    except Exception as erro:
        log_erro(f"Erro ao conectar ao arquivo Excel: {erro}")
        return None


def salvar_planilha(worksheet):
    """Salva a pasta de trabalho no disco (equivalente ao autosave da nuvem
    do Google Sheets, que não existe no Excel local)."""
    try:
        worksheet.Parent.Save()
    except Exception as erro:
        log_aviso(f"Não foi possível salvar o arquivo Excel automaticamente: {erro}")


# ==============================================================================
# LEITURA
# ==============================================================================

def obter_ultima_linha(worksheet) -> int:
    """Obtém a última linha preenchida da coluna A."""
    try:
        ultima_linha = worksheet.Cells(worksheet.Rows.Count, 1).End(XL_UP).Row
        if ultima_linha == 1 and not worksheet.Cells(1, 1).Value:
            ultima_linha = 0
        log_info(f"Última linha preenchida na coluna A: {ultima_linha}")
        return ultima_linha
    except Exception as erro:
        log_erro(f"Erro ao obter última linha: {erro}")
        return 0


def ler_dados(worksheet, linha_inicial: int):
    """
    Lê os dados da planilha, da 'linha_inicial' até a última linha
    preenchida, das colunas A até a última coluna configurada
    (`colunas.COL_ULTIMA_LEITURA`, recalculada por `aplicar_colunas()` a partir das
    colunas escolhidas no submenu "Configurar colunas" - por padrão, A:L,
    igual ao intervalo fixo que o sistema sempre usou).
    """
    if worksheet is None:
        log_erro("Worksheet inválida. Não é possível ler dados.")
        return []

    ultima_linha = obter_ultima_linha(worksheet)

    if ultima_linha < linha_inicial:
        log_aviso(
            f"A linha inicial ({linha_inicial}) é maior que a última linha "
            f"preenchida ({ultima_linha}). Nenhum dado será lido."
        )
        return []

    try:
        letra_final = colunas.indice_para_letra_coluna(colunas.COL_ULTIMA_LEITURA)
        intervalo = f"A{linha_inicial}:{letra_final}{ultima_linha}"
        valores = worksheet.Range(intervalo).Value  # tupla de tuplas (linhas x colunas)

        if valores is None:
            return []

        # Quando o intervalo tem uma única linha, o COM devolve uma tupla
        # simples em vez de uma tupla de tuplas - normaliza os dois casos.
        if not isinstance(valores[0], tuple):
            valores = (valores,)

        dados = [
            ["" if celula is None else celula for celula in linha]
            for linha in valores
        ]
        log_sucesso(f"{len(dados)} linha(s) lida(s) do intervalo {intervalo}.")
        return dados
    except Exception as erro:
        log_erro(f"Erro ao ler dados da planilha: {erro}")
        return []


def obter_celula(linha: List, indice: int) -> str:
    """
    Obtém o valor de uma célula pelo índice (base 0), retornando string vazia
    caso a linha não tenha células suficientes ou o valor seja vazio.

    Números lidos do Excel via COM vêm como `float` (ex: 18380.0), mesmo
    que a célula exiba só "18380" - quando o valor é um número inteiro,
    remove essa parte decimal espúria antes de converter para string
    (senão comandos como `.zfill(6)` não completam corretamente, já que
    "18380.0" já tem mais de 6 caracteres).
    """
    if indice < len(linha):
        valor = linha[indice]
        if valor is None:
            return ""
        if isinstance(valor, float) and valor.is_integer():
            valor = int(valor)
        return str(valor).strip()
    return ""


# ==============================================================================
# ESCRITA
# ==============================================================================
# Toda escrita no Excel passa por `executar_escrita_com_retentativa()` -
# caso o COM devolva um erro temporário (Excel ocupado processando outra
# chamada, por exemplo enquanto o usuário está com um menu aberto na
# interface), a escrita é retentada automaticamente em vez de a automação
# simplesmente falhar naquela linha.

def _e_erro_temporario_excel(erro: Exception) -> bool:
    """Detecta erros COM temporários do Excel (aplicativo ocupado), que
    costumam se resolver tentando de novo logo em seguida."""
    texto = str(erro)
    return (
        "Call was rejected by callee" in texto
        or "RPC_E_CALL_REJECTED" in texto
        or "-2147418111" in texto  # servidor COM ocupado
        or "-2146777998" in texto  # "message filter indicates..."
    )


def executar_escrita_com_retentativa(func_escrita, *args, max_tentativas: int = 6,
                                      espera_inicial: float = 1.0, **kwargs):
    """
    Executa uma escrita no Excel com retentativas e espera progressiva
    (backoff exponencial: 1s, 2s, 4s... até 10s) caso o Excel esteja
    momentaneamente ocupado (erro COM temporário).

    Qualquer outro tipo de erro (não relacionado a uma ocupação temporária)
    é repassado imediatamente, sem espera.
    """
    tentativa = 0
    espera = espera_inicial

    while True:
        try:
            return func_escrita(*args, **kwargs)
        except Exception as erro:
            if not _e_erro_temporario_excel(erro):
                raise

            tentativa += 1
            if tentativa >= max_tentativas:
                log_erro(
                    f"Excel ocupado após {tentativa} tentativa(s). Desistindo desta escrita."
                )
                raise

            log_aviso(
                f"Excel temporariamente ocupado (tentativa {tentativa}/{max_tentativas}). "
                f"Aguardando {espera:.0f}s antes de tentar novamente..."
            )
            time.sleep(espera)
            espera = min(espera * 2, 10)


def atualizar_status(worksheet, linha: int, coluna: int, status: str):
    """Atualiza uma célula específica (ex: coluna de status) com o valor informado."""
    if worksheet is None:
        log_erro("Worksheet inválida. Não é possível atualizar status.")
        return

    def _escrever():
        worksheet.Cells(linha, coluna).Value = status
        salvar_planilha(worksheet)

    try:
        executar_escrita_com_retentativa(_escrever)
        log_sucesso(f"Status da linha {linha} atualizado para '{status}'.")
    except Exception as erro:
        log_erro(f"Erro ao atualizar status da linha {linha}: {erro}")


def salvar_valor_gerado(worksheet, linha: int, coluna: int, valor: str, rotulo: str = "documento"):
    """
    Grava um identificador gerado pelo SIGEF (ex: número de CE, NL, PP ou
    OB) na planilha, na MESMA linha de onde vieram os dados que o
    originaram. Pensada para ser chamada assim que o SIGEF confirma a
    geração do documento, para que a planilha fique salva
    incrementalmente. O valor é gravado exatamente como veio do SIGEF
    (ex: "2026PP123456"), sem remover o prefixo.

    `rotulo` é usado apenas para deixar a mensagem de log mais clara (ex:
    "PP", "OB", "CE", "NL").
    """
    if worksheet is None:
        log_erro(f"Worksheet inválida. Não é possível salvar o(a) {rotulo} gerado(a).")
        return

    def _escrever():
        worksheet.Cells(linha, coluna).Value = valor
        salvar_planilha(worksheet)

    try:
        executar_escrita_com_retentativa(_escrever)
        letra_coluna = colunas.indice_para_letra_coluna(coluna)
        log_sucesso(f"{rotulo} '{valor}' salvo(a) na planilha (linha {linha}, coluna {letra_coluna}).")
    except Exception as erro:
        log_erro(f"Erro ao salvar {rotulo} '{valor}' na linha {linha}: {erro}")


def obter_proxima_coluna_livre(worksheet, linha: int, coluna_minima: Optional[int] = None) -> int:
    """
    Descobre a primeira coluna livre (vazia) na linha informada, a partir
    de `coluna_minima`, para gravar um novo valor sem sobrescrever nenhuma
    coluna já usada por outra automação.

    Quando `coluna_minima` não é informado, usa `colunas.COL_COLUNA_MINIMA` LIDA NO
    MOMENTO DA CHAMADA (e não a que existia quando o programa foi
    carregado), para respeitar a coluna de Valor configurada pelo usuário
    no submenu "Configurar colunas".
    """
    if coluna_minima is None:
        coluna_minima = colunas.COL_COLUNA_MINIMA
    try:
        if not worksheet.Cells(linha, 1).Value:
            return coluna_minima
        ultima_coluna_preenchida = worksheet.Cells(linha, worksheet.Columns.Count).End(XL_TO_LEFT).Column
        return max(ultima_coluna_preenchida + 1, coluna_minima)
    except Exception as erro:
        log_erro(f"Erro ao localizar coluna livre na linha {linha}: {erro}")
        return coluna_minima

