# Copyright (c) 2026 Pedro Henrique Carpina Farias Alves. Todos os direitos reservados.
# Software proprietário: uso, cópia, modificação e distribuição somente com
# autorização por escrito do titular. Veja o arquivo LICENSE.
"""
TEXTOS DE OBSERVAÇÃO (CE e OB)
==============================
O texto que vai no campo "Observação" da CE e da OB no SIGEF. Por padrão é
o texto de sempre; o usuário pode trocar por um texto próprio nos
Parâmetros (vale para a planilha) e nos "Parâmetros deste lançamento" do
modo Sem planilha.

No texto personalizado, duas marcações são trocadas automaticamente:
    {mes}       -> o Mês referência (ex: 06/2026)
    {processo}  -> o número do Processo
Assim o texto não precisa ser reescrito todo mês.

Na configuração, texto VAZIO significa "usar o texto padrão".
"""

TIPOS_OBSERVACAO = ("ce", "ob")

TEXTO_PADRAO_OBSERVACAO = {
    "ce": "Gratificação amigo voluntario {mes} Processo: {processo}",
    "ob": "Ressarcimento de amigos voluntario Referente ao {mes} Processo: {processo}",
}

NOMES_OBSERVACAO = {
    "ce": "Observação da CE",
    "ob": "Observação da OB",
}

# Chave em config.json (e nos parâmetros do modo Sem planilha).
CHAVES_CONFIG_OBSERVACAO = {tipo: f"observacao_{tipo}" for tipo in TIPOS_OBSERVACAO}

# Tamanho a partir do qual a tela avisa (não impede) - campos de observação
# costumam ter limite de caracteres no sistema.
LIMITE_AVISO_CARACTERES = 250


def modelo_observacao(tipo: str, config: dict) -> str:
    """O modelo em uso: o personalizado, se houver, senão o padrão."""
    personalizado = str((config or {}).get(CHAVES_CONFIG_OBSERVACAO[tipo]) or "").strip()
    return personalizado or TEXTO_PADRAO_OBSERVACAO[tipo]


def aplicar_marcacoes(modelo: str, mes_referencia: str, processo: str) -> str:
    """Troca {mes} e {processo} pelos valores. Usa troca simples de texto
    (e não .format), para que outras chaves { } escritas pelo usuário não
    causem erro - elas só ficam como estão."""
    return (
        str(modelo)
        .replace("{mes}", str(mes_referencia or ""))
        .replace("{processo}", str(processo or ""))
    )


def montar_observacao(tipo: str, config: dict) -> str:
    """Texto final que a automação escreve no SIGEF."""
    config = config or {}
    return aplicar_marcacoes(
        modelo_observacao(tipo, config), config.get("mes_referencia", ""), config.get("processo", ""),
    )
