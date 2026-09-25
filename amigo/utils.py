"""
UTILIDADES DE NEGÓCIO
=======================
Conversão e formatação de identificadores (CPF, número de documento) e de
valores monetários entre o formato da planilha e o formato esperado pelo
SIGEF. Funções puras, sem dependência de Excel/Playwright.
"""
import re
import unicodedata

from .constantes import REGEX_APENAS_DIGITOS, REGEX_PREFIXO_DOCUMENTO

def extrair_numero_documento(codigo_completo: str) -> str:
    """
    Remove o prefixo "AAAA" + sigla (ano + tipo de documento: PP, OB, CE
    ou NL) de um identificador do SIGEF, retornando apenas os dígitos
    finais que os campos de pesquisa das telas do SIGEF aceitam.

    Exemplos:
        "2026PP123456" -> "123456"
        "2026OB654321" -> "654321"
    """
    if not codigo_completo:
        return ""

    codigo = codigo_completo.strip()
    sem_prefixo = REGEX_PREFIXO_DOCUMENTO.sub("", codigo)
    return sem_prefixo or codigo


def normalizar_numero_documento(valor) -> str:
    """
    Normaliza um número de documento do SIGEF (CE, NL, OB, PP, NE) para
    comparação entre a planilha e a grade do SIGEF: remove o prefixo
    "AAAA+sigla" (reaproveitando `extrair_numero_documento`) e quaisquer
    zeros à esquerda, convertendo ambos os lados para o mesmo formato
    numérico - assim "2026CE018487" (SIGEF) e "18487" (planilha) ficam
    ambos "18487", e a comparação funciona independentemente da
    quantidade de zeros à esquerda. Nunca compara strings brutas.

    Exemplo:
        "2026CE018487" -> "18487"
        "18487"        -> "18487"
    """
    texto = str(valor).strip()

    # Defesa extra contra floats do Excel que escapem de `obter_celula`
    # (ex: "18487.0") - remove a parte decimal quando ela for só zeros.
    if "." in texto:
        parte_inteira, parte_decimal = texto.split(".", 1)
        if parte_decimal.strip("0") == "":
            texto = parte_inteira

    sem_prefixo = extrair_numero_documento(texto)
    apenas_digitos = REGEX_APENAS_DIGITOS.sub("", sem_prefixo)
    return str(int(apenas_digitos)) if apenas_digitos else ""


# Homóglifos: caracteres cirílicos/gregos VISUALMENTE idênticos a letras
# latinas. Aparecem em planilhas montadas por copiar/colar de PDF, e-mail
# ou teclado com layout trocado. O caso clássico aqui é o dígito
# verificador "X" de conta bancária vir como "Х" (U+0425, cirílico): a
# limpeza `[^0-9A-Za-z]` DESCARTA esse caractere, a conta "81.590-Х" vira
# "81590" em vez de "81590X" e nunca bate com a grade do SIGEF - embora no
# log os dois apareçam idênticos. Por isso a conversão vem ANTES da limpeza.
_HOMOGLIFOS_LATINOS = str.maketrans({
    # Cirílico maiúsculo
    "\u0410": "A", "\u0412": "B", "\u0415": "E", "\u041a": "K", "\u041c": "M",
    "\u041d": "H", "\u041e": "O", "\u0420": "P", "\u0421": "C", "\u0422": "T",
    "\u0423": "Y", "\u0425": "X", "\u0405": "S", "\u0406": "I", "\u0408": "J",
    # Cirílico minúsculo
    "\u0430": "a", "\u0432": "b", "\u0435": "e", "\u043a": "k", "\u043c": "m",
    "\u043d": "h", "\u043e": "o", "\u0440": "p", "\u0441": "c", "\u0442": "t",
    "\u0443": "y", "\u0445": "x", "\u0455": "s", "\u0456": "i", "\u0458": "j",
    # Grego maiúsculo
    "\u0391": "A", "\u0392": "B", "\u0395": "E", "\u0396": "Z", "\u0397": "H",
    "\u0399": "I", "\u039a": "K", "\u039c": "M", "\u039d": "N", "\u039f": "O",
    "\u03a1": "P", "\u03a4": "T", "\u03a5": "Y", "\u03a7": "X",
    # Grego minúsculo
    "\u03b1": "a", "\u03bd": "v", "\u03bf": "o", "\u03c1": "p", "\u03c4": "t",
    "\u03c7": "x",
})


def latinizar_homoglifos(texto: str) -> str:
    """
    Converte caracteres visualmente idênticos a letras latinas para o
    equivalente ASCII. Aplica NFKC antes (resolve largura total, como
    "\uff38" -> "X", e dígitos de outras escritas) e depois a tabela de
    homóglifos cirílicos/gregos, que o NFKC NÃO converte.
    """
    return unicodedata.normalize("NFKC", texto).translate(_HOMOGLIFOS_LATINOS)


def normalizar_chave_bancaria(valor) -> str:
    """
    Normaliza banco, agência e conta para comparação entre a planilha e a
    grade do SIGEF: mantém apenas letras e números (remove pontos,
    vírgulas, traços, barras, espaços e `&nbsp;`), passa para maiúsculo e
    tira os zeros à esquerda - assim os dois lados chegam no MESMO
    formato. É o único ponto do sistema que faz essa conversão.

    Exemplos:
        "1"        -> "1"       |  "001"        -> "1"
        "1401-X"   -> "1401X"   |  "01401X"     -> "1401X"
        "67.287-4" -> "672874"  |  "0000672874" -> "672874"
    """
    if valor is None:
        return ""

    texto = str(valor).replace("\xa0", " ")

    # Defesa contra floats do Excel que escapem de `obter_celula`
    # (ex: "74508.0" numa conta digitada sem máscara).
    if "." in texto:
        parte_inteira, parte_decimal = texto.split(".", 1)
        if parte_decimal.strip("0") == "":
            texto = parte_inteira

    limpo = re.sub(r"[^0-9A-Za-z]", "", latinizar_homoglifos(texto)).upper()
    if not limpo:
        return ""

    sem_zeros = limpo.lstrip("0")
    return sem_zeros or "0"


def normalizar_valor_sigef(valor) -> float:
    """
    Normaliza valores monetários vindos tanto da planilha (sem separador,
    ex: '55980' -> 559.80) quanto da grade do SIGEF (com separador BR, ex:
    '790,54' ou '1.234,56' -> 1234.56).
    """
    valor_str = str(valor).replace("R$", "").strip()

    if valor_str.isdigit():
        # Valor vindo da planilha (sem separador decimal).
        return float(valor_str) / 100

    # Valor vindo da grade do SIGEF.
    return float(valor_str.replace(".", "").replace(",", "."))


def formatar_cpf(valor) -> str:
    """
    Formata um CPF (vindo da planilha com ou sem máscara) no padrão
    XXX.XXX.XXX-XX exigido pelo campo #txtNuCpf do SIGEF. Caso não tenha
    exatamente 11 dígitos, devolve o valor original (sem inventar máscara).
    """
    digitos = REGEX_APENAS_DIGITOS.sub("", str(valor))
    if len(digitos) == 11:
        return f"{digitos[0:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    return str(valor).strip()


def formatar_valor_centavos(valor) -> str:
    """
    Converte o valor da planilha (ex: 1400, representando R$ 1.400,00)
    para o formato que os campos de valor em centavos do SIGEF esperam
    (#txtVlDocumento na CE, #txtValorBrutoId na NL): o mesmo número
    multiplicado por 100, sem separadores. Compartilhada entre as
    automações que preenchem valores dessa forma.

    Exemplos:
        1400        -> "140000"
        "1400"      -> "140000"
        "1.400,00"  -> "140000"
    """
    texto = str(valor).replace("R$", "").strip()
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    numero = float(texto) if texto else 0.0
    return str(int(round(numero * 100)))


def formatar_valor_br(valor) -> str:
    """
    Formata um valor monetário da planilha (ex: 1400, representando R$
    1.400,00) no padrão brasileiro com separador de milhar e 2 casas
    decimais (ex: "1.400,00") - usado para comparar com o valor líquido
    exibido na grade de resultados do SIGEF (que já vem formatado assim),
    já que a planilha guarda o valor sem separadores e o SIGEF exibe com
    separadores; sem essa conversão a comparação nunca bateria.
    """
    texto = str(valor).replace("R$", "").strip()
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    numero = float(texto) if texto else 0.0
    formatado = f"{numero:,.2f}"
    return formatado.replace(",", "X").replace(".", ",").replace("X", ".")


def obter_numero_mes_referencia(mes_referencia: str) -> str:
    """
    Extrai o número do mês (sem zero à esquerda) a partir do Mês
    Referência configurado (formato "MM/AAAA", ex: "06/2026" -> "6"),
    para selecionar a opção correta no combo de mês de competência
    (#cboMesComp) do SIGEF.
    """
    mes_str = (mes_referencia or "").split("/")[0].strip()
    return str(int(mes_str)) if mes_str.isdigit() else mes_str

