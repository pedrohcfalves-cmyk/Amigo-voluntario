"""
Reexporta os tipos/erros do Playwright usados pelo resto do sistema, com
um "stub" de fallback quando a biblioteca não está instalada - assim,
qualquer módulo pode fazer `from .playwright_compat import Page` sem
precisar repetir o try/except de import em vários lugares.
"""
try:
    from playwright.sync_api import sync_playwright, Page, BrowserContext, TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import Error as PlaywrightError
    from playwright._impl._errors import TargetClosedError
except ImportError:
    sync_playwright = None
    Page = None
    BrowserContext = None
    TargetClosedError = Exception
    PlaywrightTimeoutError = Exception
    PlaywrightError = Exception
