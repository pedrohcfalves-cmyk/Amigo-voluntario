# Amigo — Automações SIGEF

Automação de terminal para as rotinas de pagamento do **SIGEF** (Sistema
Integrado de Gestão Financeira), inspirada em fluxos de RPA (tipo Power
Automate Desktop), usando:

- **Playwright** — automação do navegador (as telas do SIGEF).
- **pywin32** — leitura e escrita no **próprio Microsoft Excel** aberto na
  tela do usuário (via COM), com cada escrita aparecendo imediatamente na
  planilha, em tempo real.

> ⚠️ **Ambiente de PRODUÇÃO.** Por padrão, o sistema aponta para o SIGEF de
> produção (`sigef.sefin.ro.gov.br`) — tudo que as automações fizerem tem
> efeito real. O único lugar que precisa ser alterado para trocar de
> ambiente é `DOMINIO_SIGEF`, em `amigo/constantes.py`.

## O que o sistema automatiza

| # | Automação | O que faz |
|---|---|---|
| 1 | **CE** | Lança uma Despesa Certificada no SIGEF. |
| 2 | **Raspar CE** | Recupera o número de uma CE já lançada. |
| 3 | **NL** | Liquida a Despesa Certificada (gera a Nota de Lançamento). |
| 4 | **Raspar NL** | Recupera o número de uma NL já gerada. |
| 5 | **PP** | Prepara o pagamento (Preparação de Pagamento). |
| 6 | **Raspar PP** | Recupera o número de uma PP já gerada. |
| 7 | **Raspar contas** | Confere banco/agência/conta gravados no SIGEF. |
| 8 | **Gerar OB** | Gera a Ordem Bancária, em lotes de 30. |
| 9 | **Raspar OB** | *(reservado — ainda não implementado)* |
| 10 | **Confirmar OB** | *(ainda não implementado)* |

As automações formam uma cadeia: **CE → NL → PP → Raspar contas → OB**.
Cada etapa lê, da planilha, o número gerado pela etapa anterior.

## Colunas Retráteis

O sistema **não exige um layout fixo de planilha**. No menu principal,
em **"Configurar colunas"**, você escolhe em qual coluna está (ou deve
ser salvo) cada dado — uma tabela visual mostra o layout atual e bloqueia
qualquer tentativa de colocar dois dados na mesma coluna, pedindo para
corrigir antes de salvar. O padrão de fábrica é:

| Coluna | Dado |
|---|---|
| B | CPF do credor/favorecido |
| D | Nota de Empenho |
| E | Banco |
| F | Agência |
| G | Conta |
| H | Valor |
| I | CE (gerada pela CE; lida pela NL/Raspar PP) |
| J | NL (gerada pela NL; lida pela PP/Raspar PP) |
| K | PP (gerada pela PP; lida pela Raspar contas/Gerar OB) |
| L | OB (gerada pela Gerar OB) |
| M | Observação da Raspar PP / Resultado da Raspar contas |
| N | Conferência da Raspar contas |

## Estrutura do projeto

```
Amigo/
├── main.py                    # ponto de entrada (python main.py)
├── config.json                 # sua configuração local (não versionado)
├── config.example.json         # modelo de config.json
├── requirements.txt
├── amigo/
│   ├── log.py                  # mensagens padronizadas no terminal
│   ├── constantes.py           # URLs do SIGEF, timeouts, regex
│   ├── playwright_compat.py    # import do Playwright com fallback
│   ├── colunas.py               # "Colunas Retráteis": layout configurável
│   ├── config_io.py             # leitura/gravação crua de config.json
│   ├── config.py                 # parâmetros gerais + menus de configuração
│   ├── excel.py                  # conectar/ler/gravar no Excel (COM)
│   ├── navegador.py              # utilidades genéricas do Playwright
│   ├── utils.py                  # formatação de CPF, valores, documentos
│   ├── menu.py                   # menu principal e execução das automações
│   └── automacoes/
│       ├── ce.py                 # CE + Raspar CE
│       ├── nl.py                 # NL + Raspar NL
│       ├── pp.py                 # PP + Raspar PP
│       ├── raspar_contas.py      # Raspar contas
│       └── ob.py                 # Gerar OB, Raspar OB, Confirmar OB
└── legado/
    └── Amigo-6.5.py              # versão anterior (arquivo único), mantida
                                   # como referência histórica
```

## Requisitos

- Windows, com o **Microsoft Excel** instalado.
- Python 3.10+.
- Acesso ao SIGEF (rede/VPN da instituição).

## Instalação

```bash
pip install -r requirements.txt
playwright install chromium
```

## Configuração

Na primeira execução, `config.json` é criado automaticamente com valores
padrão (o mesmo formato de `config.example.json`). Você também pode
copiar `config.example.json` para `config.json` e preencher à mão:

| Campo | Descrição |
|---|---|
| `data` | Data de referência das automações (ex: `30062026`). |
| `processo` | Número do processo administrativo. |
| `mes_referencia` | Mês de referência (**sempre o mês anterior**). |
| `linha_inicial` | Primeira linha da planilha a ser lida. |
| `caminho_planilha` | Caminho completo do arquivo `.xlsx`. |
| `aba_planilha` | Nome da aba (vazio = usa a 1ª). |
| `colunas` | Layout de colunas (ver "Colunas Retráteis" acima). |

## Uso

```bash
python main.py
```

O menu principal permite: configurar parâmetros, executar uma automação,
trocar de planilha, mudar a linha inicial, configurar colunas e ver as
configurações atuais.

## Uso interno

Ferramenta de uso interno para automação de rotinas financeiras. Não é
software livre nem possui licença de redistribuição.
