# Componentes de terceiros

O **Amigo - Automações SIGEF** é software proprietário (veja `LICENSE`),
mas usa os componentes de terceiros abaixo. Cada um continua regido pela
sua própria licença, e nada na licença do Amigo altera esses direitos.
Quando o Amigo é distribuído como executável (`Amigo.exe`), estes
componentes vão empacotados junto e seus avisos devem acompanhá-lo.

| Componente | Para que é usado | Licença |
|---|---|---|
| [Python](https://www.python.org/) | Linguagem e bibliotecas padrão | Python Software Foundation License |
| Tkinter / Tcl / Tk | Interface gráfica | Licença do Tcl/Tk (estilo BSD) |
| [Playwright para Python](https://github.com/microsoft/playwright-python) | Automação do navegador (telas do SIGEF) | Apache License 2.0 |
| Navegador Chromium (baixado pelo Playwright) | Navegador usado pelo Playwright | Licenças do projeto Chromium (BSD e outras) |
| [pywin32](https://github.com/mhammond/pywin32) | Controle do Microsoft Excel (COM) | Python Software Foundation License |
| [PyInstaller](https://pyinstaller.org/) (só para gerar o `.exe`) | Empacotamento do executável | GPL 2.0 com exceção de bootloader, que permite distribuir programas empacotados sob qualquer licença, inclusive proprietária |

Os textos completos das licenças estão nos sites de cada projeto,
listados acima, e nos pacotes instalados (pasta `site-packages`, arquivos
`LICENSE`/`METADATA` de cada biblioteca).

"Microsoft Excel", "Windows" e "SIGEF" são marcas ou sistemas de seus
respectivos titulares. O Amigo não é afiliado a eles; apenas os utiliza.
