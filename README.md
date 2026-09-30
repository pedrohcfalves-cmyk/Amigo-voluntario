# Amigo — Automações SIGEF  (versão 7.0)

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

## Novidades da versão 7.0

- 🗓 **Virada de ano automática.** As telas do SIGEF mudam de endereço todo
  ano (`/SIGEF2026/` → `/SIGEF2027/`). O ano agora sai da **Data** dos
  Parâmetros (Data `15012027` → SIGEF2027) e os documentos passam a sair como
  `2027CE…`, `2027NL…`, `2027PP…`, `2027OB…` sem mexer no código. Para casos
  especiais (ex: lançar em janeiro ainda no exercício anterior) existe o campo
  **"Ano do SIGEF (exercício)"** nos Parâmetros. Abas antigas do navegador de
  outro exercício não são mais reaproveitadas por engano.
- 🧪 **Modo simulação.** Marque "Modo simulação" no painel lateral: a
  automação abre o SIGEF de verdade, preenche e confere tudo, mas **para antes
  do botão final** — nada é gerado. No lugar do número, a planilha (ou a tela
  do Sem planilha) recebe `SIMULADO ✓ …` ou `SIMULADO ✗ motivo`. As etapas
  seguintes pulam linhas simuladas, e a simulação não entra nas Estatísticas.
- ⏹ **Botão Parar + andamento.** Nas automações da planilha aparece
  "linha X de Y" e o botão **"Parar depois da linha atual"** — a automação
  termina a linha em andamento e para com segurança.
- 📝 **Textos de observação personalizáveis** (CE e OB), nos Parâmetros e no
  Sem planilha: use o texto padrão ou escreva o seu, com as marcações
  `{mes}` e `{processo}` e uma prévia ao vivo.
- ✅ **Testes automáticos** (`rodar_testes.bat`), com um "SIGEF de mentira"
  que confere CE, NL, PP, OB, simulação, Parar e a virada de ano sem abrir o
  SIGEF nem o Excel.
- 🗂 Projeto versionado em **Git**. O `historico.json` (base do Relatório e
  das Estatísticas) também: depois de cada automação o programa grava um
  commit só desse arquivo, em segundo plano (sem Git instalado, ele só
  segue normalmente). Nada é enviado para a internet sozinho.

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
├── Amigo.bat                   # atalho de duplo clique (abre a interface gráfica)
├── Amigo.exe                   # aplicativo empacotado (não versionado - gerado por build_exe.bat)
├── build_exe.bat               # gera o Amigo.exe (PyInstaller), sozinho, sem perguntas
├── rodar_testes.bat            # roda os testes automáticos
├── main.py                     # ponto de entrada do modo terminal (python main.py)
├── main_gui.py                 # ponto de entrada da interface gráfica (python main_gui.py)
├── config.json                 # sua configuração local (não versionado)
├── config.example.json         # modelo de config.json
├── requirements.txt
├── tests/                      # testes automáticos (SIGEF de mentira em tests/sigef_falso.py)
└── amigo/
    ├── log.py                  # mensagens padronizadas (terminal + interface gráfica)
    ├── constantes.py           # URLs do SIGEF (com o ano do exercício), timeouts, regex
    ├── execucao.py             # modo simulação, botão Parar e andamento (linha X de Y)
    ├── observacoes.py          # textos de observação da CE e da OB (padrão ou personalizado)
    ├── playwright_compat.py    # import do Playwright com fallback
    ├── colunas.py              # "Colunas Retráteis": layout configurável
    ├── config_io.py            # leitura/gravação crua de config.json
    ├── config.py               # parâmetros gerais + menus de configuração (terminal)
    ├── setup_dependencias.py   # checa/baixa automaticamente o que faltar (interface gráfica)
    ├── gui.py                  # interface gráfica (Tkinter)
    ├── excel.py                # conectar/ler/gravar no Excel (COM)
    ├── navegador.py            # utilidades genéricas do Playwright
    ├── utils.py                # formatação de CPF, valores, documentos
    ├── relatorio.py            # histórico de execuções (historico.json) - aba "Relatório"
    ├── historico_git.py        # commit automático do historico.json no Git
    ├── estatisticas.py         # tempo economizado x tempo manual, projeção de escala - aba "Estatísticas"
    ├── lancamento_manual.py    # CE -> NL -> PP -> OB com dados digitados, sem planilha - aba "Lançamento Manual"
    ├── menu.py                 # menu principal e execução das automações (terminal)
    └── automacoes/
        ├── ce.py                 # CE + Raspar CE
        ├── nl.py                 # NL + Raspar NL
        ├── pp.py                 # PP + Raspar PP
        ├── raspar_contas.py      # Raspar contas
        └── ob.py                 # Gerar OB, Raspar OB, Confirmar OB
```

## Requisitos

- Windows, com o **Microsoft Excel** instalado.
- Python 3.10+.
- Acesso ao SIGEF (rede/VPN da instituição).

## Instalação

### Opção 1 — Executável (`Amigo.exe`)

Dê dois cliques em **`Amigo.exe`**. Não precisa ter Python instalado nem
rodar nenhum comando — o executável já leva tudo que o programa precisa
para abrir a interface gráfica. Na primeira vez que uma automação for
usada, ele mesmo baixa o navegador do Playwright, se ainda não estiver
presente.

Se `Amigo.exe` ainda não existir nesta pasta (ex: acabou de clonar o
repositório), gere-o dando dois cliques em **`build_exe.bat`** — ele
instala o PyInstaller e as dependências do projeto, empacota tudo e deixa
o `Amigo.exe` pronto nesta mesma pasta (esse passo precisa de Python
instalado só na máquina que vai *gerar* o executável; a máquina que vai
só *usar* o `Amigo.exe` pronto não precisa).

### Opção 2 — Interface gráfica via Python

Dê dois cliques em **`Amigo.bat`**. Na primeira vez, o próprio aplicativo
baixa e instala tudo que faltar (pacotes Python, navegador do Playwright)
antes de abrir a janela principal — não é preciso rodar nenhum comando.
Só é necessário ter o **Python instalado** no computador (o `Amigo.bat`
avisa e indica onde baixar, caso não encontre).

### Opção 3 — Terminal / manual

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
| `ano_sigef` | Ano do exercício do SIGEF (vazio = usa o ano da `data`). |
| `observacao_ce` | Texto de observação da CE (vazio = texto padrão). Aceita `{mes}` e `{processo}`. |
| `observacao_ob` | Texto de observação da OB (vazio = texto padrão). Aceita `{mes}` e `{processo}`. |

## Uso

**Interface gráfica** (janela com botões, recomendado para o dia a dia):

```bash
python main_gui.py
```
ou dois cliques em `Amigo.bat`. A janela abre na tela **Início** (escolha entre
lançar com ou sem planilha) e tem um menu lateral com as telas — **Automações**
(conectar a planilha e rodar cada etapa), **Lançamento Manual** (ver
abaixo), **Parâmetros** (data, processo,
mês de referência, linha inicial, arquivo Excel), **Colunas Retráteis**
(a mesma tabela do modo terminal, com os campos em conflito destacados em
vermelho e o botão "Salvar colunas" bloqueado até corrigir), **Estatísticas**
(tempo economizado com a automação frente a uma estimativa de tempo manual
ajustável, e uma projeção de tempo para volumes maiores de itens — dados
para embasar um estudo de implementação em larga escala) e **Relatório**
(quantas CE/NL/PP/OB foram feitas por mês) — mais um painel de log
embaixo, sempre visível.

### Lançamento Manual (sem planilha)

Alternativa à planilha: na aba **Lançamento Manual** o usuário digita, para
cada pessoa, CPF, Nota de Empenho, Banco, Agência, Conta e Valor (ex:
`1400` = R$ 1.400,00). Cada campo é conferido e padronizado na hora (CPF
com dígito verificador, NE no formato do SIGEF, valor em reais etc.), e o
programa faz **CE → NL → PP → OB, uma pessoa de cada vez**: a CE gerada é
usada na NL, a NL na PP e a PP na OB (1 OB por pessoa).

- **Nada é salvo**: nem na planilha, nem em arquivo, nem no histórico do
  Relatório/Estatísticas. Os 4 números aparecem na tabela "Resultados", com
  o botão "Copiar resultados" (cola direto no Excel/Word).
- Se uma pessoa parar em alguma etapa, o motivo aparece ao clicar na linha,
  e "Tentar de novo as que pararam" continua da etapa em que parou, sem
  refazer (nem duplicar no SIGEF) as anteriores.
- Usa exatamente as mesmas automações da planilha, com uma "planilha em
  memória" (`lancamento_manual.PlanilhaVirtual`) — a automação normal da
  planilha não muda. Só 1 automação usa o navegador do SIGEF por vez: enquanto
  uma roda, os botões da outra ficam bloqueados.
- Tem **parâmetros próprios** (Data, Processo e Mês referência, no mesmo
  formato da aba **Parâmetros**), guardados à parte em `config.json`
  (chave `parametros_lancamento_manual`) — não alteram os da planilha.

**Modo terminal** (equivalente, por menus numerados):

```bash
python main.py
```

Em ambos os modos: configurar parâmetros, executar uma automação, trocar
de planilha, mudar a linha inicial, configurar colunas e ver as
configurações atuais. As duas interfaces leem e gravam o mesmo
`config.json` e chamam exatamente as mesmas automações — é só uma questão
de preferência.

## Testes automáticos

Dê dois cliques em `rodar_testes.bat` (ou rode
`python -m unittest discover -s tests -t . -v`). Os testes **não** abrem o
SIGEF nem o Excel: usam um SIGEF de mentira (`tests/sigef_falso.py`) e uma
planilha em memória, e conferem as automações CE/NL/PP/OB (normal e
simulação), o botão Parar, o andamento, a virada de ano, os textos de
observação, a padronização dos dados, a cadeia do Sem planilha, as
Estatísticas e as Colunas Retráteis.

## Licença

Copyright (c) 2026 **Pedro Henrique Carpina Farias Alves**. Todos os
direitos reservados.

Software **proprietário**: é proibido copiar, modificar, distribuir,
publicar ou usar este programa sem autorização prévia e por escrito do
titular. Os termos completos estão em [`LICENSE`](LICENSE).

As bibliotecas de terceiros usadas pelo programa (Python, Tkinter,
Playwright, pywin32, PyInstaller) seguem as próprias licenças - veja
[`THIRD-PARTY-NOTICES.md`](THIRD-PARTY-NOTICES.md).
