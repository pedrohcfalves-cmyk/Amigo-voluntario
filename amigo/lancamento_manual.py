# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
LANÇAMENTO MANUAL (sem planilha)
================================
Alternativa à leitura da planilha do Excel: o usuário DIGITA, numa aba
própria da interface gráfica, os dados de cada pessoa (CPF, Nota de
Empenho, Banco, Agência, Conta e Valor), e o programa faz a cadeia
completa CE -> NL -> PP -> OB para UMA pessoa de cada vez:

    CE gerada  -> usada para fazer a NL
    NL gerada  -> usada para fazer a PP
    PP gerada  -> usada para fazer a OB

No final, os 4 números (CE, NL, PP, OB) aparecem na tela. NADA é gravado
em lugar nenhum: nem na planilha do Excel, nem em arquivo, nem no
histórico do Relatório/Estatísticas.

Como isso funciona SEM alterar nenhuma automação: as automações
(amigo.automacoes) recebem `(dados, config, worksheet)` - uma lista de
linhas e uma planilha onde gravam o resultado. Aqui, cada pessoa vira
UMA linha montada em memória (nas mesmas posições das "Colunas
Retráteis" configuradas) e a "planilha" é uma `PlanilhaVirtual`: um
objeto que se comporta como a planilha do Excel para as automações
(`worksheet.Cells(linha, coluna).Value = ...` e `worksheet.Parent.Save()`),
mas só guarda os valores na memória do programa - e o "Save" não faz
nada, de propósito. Assim as automações rodam exatamente o mesmo código
da planilha, com as mesmas conferências, e a automação normal da
planilha continua intocada.

Este módulo não conhece nada de Tkinter - a aba "Lançamento Manual"
(amigo/gui.py) só chama as funções daqui.
"""
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Callable, Dict, List, Optional, Tuple

from . import colunas
from .constantes import REGEX_DOCUMENTO_OB, REGEX_DOCUMENTO_PP
from .execucao import eh_simulado
from .log import adicionar_ouvinte, log_erro, log_info, log_sucesso, remover_ouvinte
from .observacoes import CHAVES_CONFIG_OBSERVACAO
from .utils import latinizar_homoglifos

# Os 6 dados que o usuário digita, na ordem em que aparecem na tela.
CAMPOS_LANCAMENTO = ["cpf", "ne", "banco", "agencia", "conta", "valor"]

# As 4 etapas da cadeia, na ordem em que rodam, com o nome "amigável"
# mostrado para o usuário.
ETAPAS = ["CE", "NL", "PP", "OB"]
NOMES_ETAPAS = {
    "CE": "CE - Despesa Certificada",
    "NL": "NL - Nota de Lançamento",
    "PP": "PP - Preparação de Pagamento",
    "OB": "OB - Ordem Bancária",
}

# Parâmetros da aba "Parâmetros" sem os quais a cadeia não pode rodar
# (usados pelas próprias automações CE/NL/PP/OB).
PARAMETROS_OBRIGATORIOS = [
    ("data", "Data"),
    ("processo", "Processo"),
    ("mes_referencia", "Mês referência"),
]


# ==============================================================================
# PADRONIZAÇÃO DOS DADOS DIGITADOS
# ==============================================================================
# Cada função recebe o texto como o usuário digitou e devolve
# (valor_padronizado, texto_para_mostrar) - ou levanta ValueError com uma
# mensagem em português simples, pronta para aparecer na tela, dizendo o
# que corrigir. O "valor padronizado" é exatamente o formato que a
# automação já espera encontrar numa célula da planilha.

def _limpar(texto) -> str:
    """Tira espaços (inclusive o 'espaço duro' de copiar/colar) e troca
    letras de outros alfabetos que parecem latinas (ver
    `utils.latinizar_homoglifos`)."""
    texto = latinizar_homoglifos(str(texto or "")).replace("\xa0", " ")
    return re.sub(r"\s+", "", texto).upper()


def _cpf_valido(digitos: str) -> bool:
    """Confere os 2 dígitos verificadores do CPF - pega a maioria dos
    erros de digitação (um número trocado ou dois números invertidos)."""
    if len(digitos) != 11 or digitos == digitos[0] * 11:
        return False
    for tamanho in (9, 10):
        soma = sum(int(digitos[i]) * (tamanho + 1 - i) for i in range(tamanho))
        digito = (soma * 10) % 11 % 10
        if digito != int(digitos[tamanho]):
            return False
    return True


def padronizar_cpf(texto) -> Tuple[str, str]:
    digitos = re.sub(r"\D", "", _limpar(texto))
    if not digitos:
        raise ValueError("Digite o CPF.")
    if len(digitos) != 11:
        raise ValueError(f"O CPF precisa ter 11 números - você digitou {len(digitos)}.")
    if not _cpf_valido(digitos):
        raise ValueError("Este CPF não existe. Confira se algum número foi digitado errado.")
    exibicao = f"{digitos[0:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    return digitos, exibicao


def padronizar_nota_empenho(texto) -> Tuple[str, str]:
    """
    Aceita "2026NE001234", "2026 NE 1234", "NE1234" ou só "1234".
    Guarda no formato do SIGEF: "AAAANE" + 6 números quando o ano foi
    informado, ou só os 6 números - a automação NL usa apenas os números
    (`utils.extrair_numero_documento`) para pesquisar o empenho.
    """
    limpo = _limpar(texto).replace(".", "").replace("-", "").replace("/", "")
    if not limpo:
        raise ValueError("Digite o número da Nota de Empenho.")
    busca = re.fullmatch(r"(?:(\d{4})NE|NE)?(\d+)", limpo)
    if not busca:
        raise ValueError("Use só números, com ou sem o ano e o 'NE' na frente (ex: 2026NE001234 ou 1234).")
    ano, numero = busca.group(1), busca.group(2)
    if len(numero) > 6:
        raise ValueError(
            "O número do empenho tem mais de 6 números. Se você digitou o ano junto, "
            "coloque 'NE' entre o ano e o número (ex: 2026NE001234)."
        )
    numero = numero.zfill(6)
    if int(numero) == 0:
        raise ValueError("O número do empenho não pode ser zero.")
    valor = f"{ano}NE{numero}" if ano else numero
    return valor, valor


def padronizar_banco(texto) -> Tuple[str, str]:
    limpo = _limpar(texto)
    if not limpo:
        raise ValueError("Digite o número do banco.")
    if not limpo.isdigit():
        raise ValueError("O banco é só números (ex: 1 para o Banco do Brasil).")
    sem_zeros = limpo.lstrip("0")
    if not sem_zeros:
        raise ValueError("O banco não pode ser zero.")
    if len(sem_zeros) > 3:
        raise ValueError("O número do banco tem no máximo 3 números (ex: 1, 104, 237).")
    return sem_zeros, sem_zeros.zfill(3)


def padronizar_agencia(texto) -> Tuple[str, str]:
    limpo = _limpar(texto).replace(".", "")
    if not limpo:
        raise ValueError("Digite a agência.")
    if not re.fullmatch(r"\d{1,6}(-?[0-9X])?", limpo):
        raise ValueError("A agência é só números, com ou sem o dígito no final (ex: 1178-9 ou 2757X).")
    return limpo, limpo


def padronizar_conta(texto) -> Tuple[str, str]:
    limpo = _limpar(texto)
    sem_pontos = limpo.replace(".", "")
    if not sem_pontos:
        raise ValueError("Digite a conta.")
    if not re.fullmatch(r"\d{1,15}(-?[0-9X])?", sem_pontos):
        raise ValueError("A conta é só números, com ou sem o dígito no final (ex: 74.508-1).")
    if not re.search(r"[1-9]", sem_pontos):
        raise ValueError("A conta não pode ser zero.")
    return limpo, limpo


def _formatar_reais(valor: Decimal) -> str:
    texto = f"{valor:,.2f}"  # 1,400.00
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


def padronizar_valor(texto) -> Tuple[str, str]:
    """
    Aceita o valor como a pessoa costuma escrever e devolve sempre no
    formato "1400,00" (reais, com vírgula) - que é lido do MESMO jeito por
    todas as automações (`utils.formatar_valor_centavos`,
    `utils.formatar_valor_br`, `pp._valor_planilha_para_float`):

        "1400"        -> R$ 1.400,00   (número sem vírgula = reais inteiros)
        "1400,50"     -> R$ 1.400,50
        "1.400,50"    -> R$ 1.400,50
        "R$ 1.400"    -> R$ 1.400,00   (ponto seguido de 3 números = milhar)
        "1400.50"     -> R$ 1.400,50   (ponto seguido de 1 ou 2 números = centavos)
    """
    limpo = _limpar(texto).replace("R$", "")
    if not limpo:
        raise ValueError("Digite o valor.")

    if "," in limpo:
        if limpo.count(",") > 1 or not re.fullmatch(r"\d{1,3}(\.\d{3})*,\d{1,2}|\d+,\d{1,2}", limpo):
            raise ValueError("Valor com formato estranho. Exemplos certos: 1400 ou 1400,50 ou 1.400,50.")
        numero = limpo.replace(".", "").replace(",", ".")
    elif "." in limpo:
        if re.fullmatch(r"\d{1,3}(\.\d{3})+", limpo):
            numero = limpo.replace(".", "")
        elif re.fullmatch(r"\d+\.\d{1,2}", limpo):
            numero = limpo
        else:
            raise ValueError("Valor com formato estranho. Exemplos certos: 1400 ou 1400,50 ou 1.400,50.")
    elif limpo.isdigit():
        numero = limpo
    else:
        raise ValueError("O valor é só números (ex: 1400 para R$ 1.400,00).")

    try:
        valor = Decimal(numero).quantize(Decimal("0.01"))
    except InvalidOperation:
        raise ValueError("Não entendi o valor. Exemplo: 1400 para R$ 1.400,00.")
    if valor <= 0:
        raise ValueError("O valor precisa ser maior que zero.")

    reais, centavos = divmod(int(valor * 100), 100)
    return f"{reais},{centavos:02d}", _formatar_reais(valor)


PADRONIZADORES: Dict[str, Callable[[str], Tuple[str, str]]] = {
    "cpf": padronizar_cpf,
    "ne": padronizar_nota_empenho,
    "banco": padronizar_banco,
    "agencia": padronizar_agencia,
    "conta": padronizar_conta,
    "valor": padronizar_valor,
}


def padronizar_campo(campo: str, texto) -> Tuple[str, str]:
    """Padroniza 1 campo (ver `PADRONIZADORES`). Levanta ValueError com a
    explicação do que corrigir."""
    return PADRONIZADORES[campo](texto)


def montar_item(textos: Dict[str, str]) -> Tuple[Optional[dict], Dict[str, str]]:
    """
    Padroniza os 6 campos digitados de uma vez. Devolve (item, erros):
    `item` é None se algum campo tiver erro; `erros` é {campo: mensagem}.
    O item guarda o valor padronizado de cada campo e, em "exibicao", o
    texto bonito para mostrar na tela (ex: CPF com pontos, R$ no valor).
    """
    item = {"exibicao": {}}
    erros = {}
    for campo in CAMPOS_LANCAMENTO:
        try:
            valor, exibicao = padronizar_campo(campo, textos.get(campo, ""))
            item[campo] = valor
            item["exibicao"][campo] = exibicao
        except ValueError as erro:
            erros[campo] = str(erro)
    if erros:
        return None, erros
    for etapa in ETAPAS:
        item[etapa.lower()] = ""  # resultados, preenchidos pela cadeia
    item.update({"situacao": "pendente", "etapa_parada": "", "motivo": ""})
    return item, {}


def parametros_faltando(config: dict) -> List[str]:
    """Nomes (amigáveis) dos parâmetros obrigatórios ainda vazios."""
    return [rotulo for chave, rotulo in PARAMETROS_OBRIGATORIOS if not str(config.get(chave) or "").strip()]


# ==============================================================================
# PARÂMETROS PRÓPRIOS DO LANÇAMENTO MANUAL
# ==============================================================================
# Os mesmos 3 parâmetros da aba "Parâmetros" (Data, Processo, Mês
# referência), no MESMO formato que as automações usam - mas SEPARADOS:
# ficam em config.json na chave abaixo, e nunca mexem nos parâmetros da
# automação da planilha ("data", "processo", "mes_referencia").
CHAVE_CONFIG_PARAMETROS_MANUAL = "parametros_lancamento_manual"


def padronizar_data(texto) -> Tuple[str, str]:
    """
    Aceita "30062026", "30/06/2026" ou "30-06-2026" e devolve "30062026" -
    o mesmo formato do parâmetro "Data" da automação normal (é o que os
    campos de data do SIGEF recebem).
    """
    digitos = re.sub(r"\D", "", str(texto or ""))
    if not digitos:
        raise ValueError("Digite a data.")
    if len(digitos) != 8:
        raise ValueError("A data precisa ter dia, mês e ano com 4 números (ex: 30062026 ou 30/06/2026).")
    dia, mes, ano = int(digitos[:2]), int(digitos[2:4]), int(digitos[4:])
    try:
        datetime(ano, mes, dia)
    except ValueError:
        raise ValueError("Essa data não existe. Confira o dia e o mês (ex: 30062026 = 30/06/2026).")
    if not 2000 <= ano <= 2100:
        raise ValueError("Confira o ano da data (ex: 2026).")
    return digitos, f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}"


def padronizar_processo(texto) -> Tuple[str, str]:
    """O processo é usado como texto na observação da CE e da OB -
    mesmo formato livre da automação normal; só tira espaços das pontas."""
    processo = re.sub(r"\s+", "", str(texto or ""))
    if not processo:
        raise ValueError("Digite o número do processo.")
    if not re.search(r"\d", processo):
        raise ValueError("O processo precisa ter números (ex: 0029.037004/2026-72).")
    return processo, processo


def padronizar_mes_referencia(texto) -> Tuple[str, str]:
    """Aceita "06/2026", "6/2026" ou "062026" e devolve "06/2026" - o
    formato "MM/AAAA" que a automação normal usa."""
    limpo = re.sub(r"\s+", "", str(texto or ""))
    if not limpo:
        raise ValueError("Digite o mês referência.")
    busca = re.fullmatch(r"(\d{1,2})[/.-]?(\d{4})", limpo)
    if not busca:
        raise ValueError("Use mês e ano (ex: 06/2026).")
    mes, ano = int(busca.group(1)), int(busca.group(2))
    if not 1 <= mes <= 12:
        raise ValueError("O mês vai de 01 a 12.")
    if not 2000 <= ano <= 2100:
        raise ValueError("Confira o ano (ex: 2026).")
    valor = f"{mes:02d}/{ano}"
    return valor, valor


PADRONIZADORES_PARAMETROS: Dict[str, Callable[[str], Tuple[str, str]]] = {
    "data": padronizar_data,
    "processo": padronizar_processo,
    "mes_referencia": padronizar_mes_referencia,
}


def padronizar_parametro_manual(chave: str, texto) -> Tuple[str, str]:
    """Padroniza 1 parâmetro do Lançamento Manual (ver
    `PADRONIZADORES_PARAMETROS`). Levanta ValueError com o que corrigir."""
    return PADRONIZADORES_PARAMETROS[chave](texto)


def aviso_mes_referencia(data: str, mes_referencia: str) -> str:
    """
    Lembrete (não bloqueia): o mês referência costuma ser o mês ANTERIOR
    ao da Data - a mesma regra lembrada na aba "Parâmetros". Devolve o
    texto do aviso, ou "" se estiver tudo de acordo (ou sem como conferir).
    """
    try:
        data_ok, _ = padronizar_data(data)
        mes_ok, _ = padronizar_mes_referencia(mes_referencia)
    except ValueError:
        return ""
    mes_data, ano_data = int(data_ok[2:4]), int(data_ok[4:])
    mes_anterior = f"{12:02d}/{ano_data - 1}" if mes_data == 1 else f"{mes_data - 1:02d}/{ano_data}"
    if mes_ok != mes_anterior:
        return f"Atenção: normalmente é o mês anterior ao da Data ({mes_anterior}). Confira se está certo."
    return ""


def parametros_manuais_salvos(config: dict) -> Dict[str, str]:
    """Parâmetros do Lançamento Manual guardados em config.json (vazios na
    1ª vez)."""
    salvos = config.get(CHAVE_CONFIG_PARAMETROS_MANUAL)
    if not isinstance(salvos, dict):
        salvos = {}
    chaves = list(PADRONIZADORES_PARAMETROS) + list(CHAVES_CONFIG_OBSERVACAO.values())
    return {chave: str(salvos.get(chave) or "") for chave in chaves}


def montar_parametros_manuais(textos: Dict[str, str]) -> Tuple[Optional[Dict[str, str]], Dict[str, str]]:
    """Padroniza os 3 parâmetros digitados na aba. Devolve (parametros,
    erros) - `parametros` é None se algum tiver erro."""
    parametros, erros = {}, {}
    for chave, padronizar in PADRONIZADORES_PARAMETROS.items():
        try:
            parametros[chave], _ = padronizar(textos.get(chave, ""))
        except ValueError as erro:
            erros[chave] = str(erro)
    # Textos de observação: vazio = texto padrão (ver amigo.observacoes).
    for chave in CHAVES_CONFIG_OBSERVACAO.values():
        parametros[chave] = str(textos.get(chave) or "").strip()
    return (None, erros) if erros else (parametros, {})


def config_para_lancamento_manual(config: dict, parametros: Dict[str, str]) -> dict:
    """
    Cópia da configuração para as automações do Lançamento Manual: tudo
    igual à configuração geral (colunas, conta de origem da OB etc.),
    MENOS Data, Processo e Mês referência, que vêm dos parâmetros
    próprios desta aba. A configuração original não é alterada.
    """
    copia = dict(config)
    copia.update(parametros)
    return copia


# ==============================================================================
# "PLANILHA" EM MEMÓRIA
# ==============================================================================

class _CelulaVirtual:
    def __init__(self, valores: dict, linha: int, coluna: int):
        self._valores = valores
        self._chave = (linha, coluna)

    @property
    def Value(self):  # mesmo nome da propriedade COM do Excel
        return self._valores.get(self._chave)

    @Value.setter
    def Value(self, valor):
        self._valores[self._chave] = valor


class _PastaVirtual:
    def Save(self):
        """De propósito, não salva nada em lugar nenhum."""


class PlanilhaVirtual:
    """
    Imita o pouco da planilha do Excel (objeto COM) que as automações usam
    para GRAVAR resultados - `excel.salvar_valor_gerado` e
    `excel.atualizar_status` fazem `worksheet.Cells(l, c).Value = valor` e
    `worksheet.Parent.Save()`. Os valores ficam só na memória, enquanto o
    programa estiver aberto.
    """
    Name = "Lançamento manual (só na memória)"
    # Lido por `excel._somente_memoria` para as mensagens do painel não
    # dizerem que o número foi salvo "na planilha".
    SOMENTE_MEMORIA = True

    def __init__(self):
        self._valores = {}
        self.Parent = _PastaVirtual()

    def Cells(self, linha: int, coluna: int) -> _CelulaVirtual:
        return _CelulaVirtual(self._valores, linha, coluna)

    def valor(self, linha: int, coluna: int) -> str:
        valor = self._valores.get((linha, coluna))
        return "" if valor is None else str(valor).strip()


def _montar_linha(item: dict) -> list:
    """
    Monta a "linha da planilha" de 1 pessoa, com cada dado na coluna em
    que a automação vai procurar - sempre as colunas configuradas AGORA em
    "Colunas Retráteis" (lidas de `colunas.COL_*` na hora, nunca fixas).
    Os documentos já gerados (quando a pessoa está continuando de onde
    parou) entram nas mesmas colunas que a etapa seguinte lê.
    """
    linha = [""] * colunas.COL_ULTIMA_LEITURA
    linha[colunas.COL_CE_CPF] = item["cpf"]
    linha[colunas.COL_NL_NE] = item["ne"]
    linha[colunas.COL_PP_BANCO] = item["banco"]
    linha[colunas.COL_PP_AGENCIA] = item["agencia"]
    linha[colunas.COL_PP_CONTA] = item["conta"]
    linha[colunas.COL_CE_VALOR] = item["valor"]
    if item.get("ce"):
        linha[colunas.COL_NL_CE] = item["ce"]          # a NL e a PP leem a CE daqui
    if item.get("nl"):
        linha[colunas.COL_PP_NL] = item["nl"]          # a PP lê a NL daqui
    if item.get("pp"):
        linha[colunas.COL_OB_PP_ESPERADO] = item["pp"]  # a OB lê a PP daqui
    return linha


# ==============================================================================
# A CADEIA CE -> NL -> PP -> OB (1 pessoa)
# ==============================================================================

def _funcoes_das_etapas():
    """Import tardio: as automações puxam o Playwright, que só é preciso
    quando o usuário realmente manda rodar (e assim os testes deste
    módulo conseguem trocar as funções por simulações)."""
    from .automacoes import ce, gerar_ob, nl, pp
    return {"CE": ce, "NL": nl, "PP": pp, "OB": gerar_ob}


def _coluna_de_saida(etapa: str) -> int:
    return {
        "CE": colunas.COL_CE_GERADA,
        "NL": colunas.COL_NL_GERADA,
        "PP": colunas.COL_PP_GERADA,
        "OB": colunas.COL_OB_GERADA,
    }[etapa]


def _etapa_deu_certo(etapa: str, retorno, valor: str) -> bool:
    """Cada automação devolve quantos itens processou; a PP e a OB, além
    disso, gravam na MESMA célula a mensagem de erro do SIGEF quando
    falham - por isso o valor gravado também é conferido."""
    if not valor:
        return False
    if etapa == "OB":
        return isinstance(retorno, tuple) and retorno[0] >= 1 and bool(REGEX_DOCUMENTO_OB.fullmatch(valor))
    processados = retorno if isinstance(retorno, int) else 0
    if processados < 1:
        return False
    if etapa == "PP":
        return bool(REGEX_DOCUMENTO_PP.fullmatch(valor))
    return True


def _rodar_etapa(etapa: str, funcao, item: dict, numero_item: int, config: dict) -> Tuple[bool, str, str]:
    """
    Roda 1 automação (CE, NL, PP ou OB) para 1 pessoa. Devolve
    (deu_certo, numero_gerado, motivo_se_falhou). O motivo vem da própria
    mensagem do SIGEF gravada na célula (PP/OB) ou da última mensagem de
    erro/aviso que a automação escreveu no painel de mensagens.
    """
    planilha = PlanilhaVirtual()
    config_item = dict(config)
    # As automações numeram a "linha" como linha_inicial + posição; com 1
    # linha só, isso faz as mensagens do painel dizerem "Linha N" = pessoa N.
    config_item["linha_inicial"] = numero_item

    mensagens_problema = []

    def ouvinte(nivel: str, mensagem: str):
        if nivel in ("erro", "aviso"):
            mensagens_problema.append(mensagem)

    erro_inesperado = None
    adicionar_ouvinte(ouvinte)
    try:
        retorno = funcao([_montar_linha(item)], config_item, planilha)
    except Exception as erro:  # a automação já trata os erros por linha; isto é só a rede de segurança
        retorno = None
        erro_inesperado = erro
    finally:
        remover_ouvinte(ouvinte)

    if erro_inesperado is not None:
        mensagens_problema.append(f"Erro inesperado: {erro_inesperado}")
        log_erro(f"Pessoa {numero_item}: erro inesperado na etapa {etapa}: {erro_inesperado}")

    valor = planilha.valor(numero_item, _coluna_de_saida(etapa))
    # Modo simulação: "SIMULADO ✓ ..." = tudo certo até antes de confirmar;
    # "SIMULADO ✗ motivo" = o que daria errado.
    if eh_simulado(valor):
        if "✓" in valor[:12]:
            return "simulado", valor, ""
        return False, "", valor
    if _etapa_deu_certo(etapa, retorno, valor):
        return True, valor, ""

    if valor:
        motivo = valor  # PP/OB: a mensagem do SIGEF foi gravada no lugar do número
    elif mensagens_problema:
        motivo = mensagens_problema[-1]
    else:
        motivo = f"A etapa {etapa} terminou sem gerar o número (veja as mensagens do sistema)."
    return False, "", motivo


def executar_cadeia(
    item: dict,
    numero_item: int,
    config: dict,
    ao_iniciar_etapa: Optional[Callable[[str], None]] = None,
    ao_terminar_etapa: Optional[Callable[[str, bool, str], None]] = None,
    funcoes: Optional[dict] = None,
) -> dict:
    """
    Faz CE -> NL -> PP -> OB para 1 pessoa, uma etapa depois da outra,
    passando o número gerado por uma etapa para a seguinte. Para na
    primeira etapa que falhar (não adianta tentar a NL sem CE, etc.).

    Etapas que o `item` JÁ tem (ex: a pessoa parou na NL numa tentativa
    anterior e já tem a CE) são puladas - assim "tentar de novo" continua
    de onde parou, sem gerar uma CE repetida no SIGEF.

    Atualiza e devolve o próprio `item` (campos "ce", "nl", "pp", "ob",
    "situacao" = "concluido"/"parou", "etapa_parada", "motivo").
    `ao_iniciar_etapa(etapa)` e `ao_terminar_etapa(etapa, ok, texto)` são
    avisos para a tela acompanhar o andamento.
    """
    funcoes = funcoes or _funcoes_das_etapas()
    cpf = item["exibicao"]["cpf"]
    item.update({"situacao": "andamento", "etapa_parada": "", "motivo": ""})

    for etapa in ETAPAS:
        chave = etapa.lower()
        if item.get(chave):
            log_info(f"Pessoa {numero_item} (CPF {cpf}): {etapa} já existe ({item[chave]}) - pulando esta etapa.")
            if ao_terminar_etapa:
                ao_terminar_etapa(etapa, True, item[chave])
            continue

        log_info(f"Pessoa {numero_item} (CPF {cpf}): começando a etapa {NOMES_ETAPAS[etapa]}...")
        if ao_iniciar_etapa:
            ao_iniciar_etapa(etapa)

        ok, numero, motivo = _rodar_etapa(etapa, funcoes[etapa], item, numero_item, config)
        if ao_terminar_etapa:
            ao_terminar_etapa(etapa, bool(ok), numero if ok else motivo)

        if ok == "simulado":
            # A etapa seguinte precisaria do número de verdade (a NL pesquisa
            # a CE no SIGEF, a PP pesquisa a NL...) - numa simulação ele não
            # existe, então a simulação desta pessoa termina aqui.
            indice = ETAPAS.index(etapa)
            proxima = ETAPAS[indice + 1] if indice + 1 < len(ETAPAS) else ""
            item[chave] = numero
            motivo_simulado = f"Simulação: a {etapa} foi conferida e parou antes de confirmar."
            if proxima:
                motivo_simulado += (
                    f" A {proxima} precisa de uma {etapa} de verdade, por isso a simulação para aqui."
                )
            item.update({"situacao": "simulado", "etapa_parada": proxima, "motivo": motivo_simulado})
            log_sucesso(f"Pessoa {numero_item} (CPF {cpf}): {motivo_simulado}")
            return item

        if not ok:
            item.update({"situacao": "parou", "etapa_parada": etapa, "motivo": motivo})
            log_erro(f"Pessoa {numero_item} (CPF {cpf}): parou na etapa {etapa}. Motivo: {motivo}")
            return item

        item[chave] = numero
        log_sucesso(f"Pessoa {numero_item} (CPF {cpf}): {etapa} {numero} gerada.")

    item["situacao"] = "concluido"
    log_sucesso(
        f"Pessoa {numero_item} (CPF {cpf}) concluída: CE {item['ce']} | NL {item['nl']} | "
        f"PP {item['pp']} | OB {item['ob']}."
    )
    return item


def texto_resultados(itens: List[Tuple[int, dict]]) -> str:
    """Texto separado por TAB (cola direto no Excel ou no bloco de notas)
    com os resultados - usado pelo botão "Copiar resultados"."""
    linhas = ["Nº\tCPF\tNota de Empenho\tValor\tCE\tNL\tPP\tOB\tSituação"]
    for numero, item in itens:
        if item["situacao"] == "concluido":
            situacao = "Concluído"
        elif item["situacao"] == "simulado":
            situacao = "Simulado (nada foi gerado)"
        else:
            situacao = f"Parou na {item['etapa_parada']}: {item['motivo']}"
        linhas.append("\t".join([
            str(numero), item["exibicao"]["cpf"], item["exibicao"]["ne"], item["exibicao"]["valor"],
            item["ce"] or "-", item["nl"] or "-", item["pp"] or "-", item["ob"] or "-", situacao,
        ]))
    return "\n".join(linhas)
