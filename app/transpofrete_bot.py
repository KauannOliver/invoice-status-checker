from __future__ import annotations

import asyncio
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Callable

from playwright.async_api import (
    Browser,
    BrowserContext,
    Locator,
    Page,
    Route,
    TimeoutError as PlaywrightTimeoutError,
    async_playwright,
)

from .config import EXCEL, PORTAL, SELECTORS
from .models import InvoiceResult, InvoiceTask

LogFn = Callable[[str], None]
ProgressFn = Callable[[int, int], None]


INVOICE_FILTER_SELECTORS = (
    'th:has(span.ui-column-title:has-text("Número")) input.ui-column-filter',
    'th:has(span.ui-column-title:has-text("Numero")) input.ui-column-filter',
    'input[id*="listaTable"][id$=":filter"]',
    'thead input.ui-column-filter',
)

ROWS_PER_PAGE_SELECTOR = 'select[name$="_rppDD"]'

STATUS_LABELS = (
    "Gerada",
    "Aceita",
    "Rejeitada",
    "Conferida",
    "Em Geração",
    "Em Geracao",
    "Paga",
)


@dataclass
class BotOptions:
    username: str
    password: str
    headless: bool = False
    max_tabs: int = 5
    slow_mo_ms: int = 0
    timeout_ms: int = 45000
    fast_mode: bool = True
    settle_ms: int = 900


async def _safe_log(log: LogFn | None, message: str) -> None:
    if log:
        log(message)


async def _settle(page: Page, options: BotOptions | None = None, ms: int | None = None) -> None:
    delay = ms if ms is not None else (options.settle_ms if options else 800)

    if delay > 0:
        await page.wait_for_timeout(delay)


def _format_from(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y 00:00:00")


async def _wait_ajax(page: Page, timeout_ms: int = 15000) -> None:
    try:
        await page.wait_for_load_state("domcontentloaded", timeout=timeout_ms)
    except PlaywrightTimeoutError:
        pass

    try:
        await page.wait_for_function(
            """
            () => {
                const jqueryOk = !window.jQuery || window.jQuery.active === 0;

                const primefacesOk =
                    !window.PrimeFaces ||
                    !PrimeFaces.ajax ||
                    !PrimeFaces.ajax.Queue ||
                    PrimeFaces.ajax.Queue.isEmpty();

                return jqueryOk && primefacesOk;
            }
            """,
            timeout=timeout_ms,
        )
    except PlaywrightTimeoutError:
        pass


async def _is_login_visible(page: Page) -> bool:
    try:
        return await page.locator(SELECTORS.login_input).is_visible(timeout=2000)
    except Exception:
        return False


async def _click_locator(locator: Locator, timeout_ms: int, settle_ms: int = 600) -> bool:
    try:
        await locator.wait_for(state="visible", timeout=timeout_ms)
        await locator.click(timeout=timeout_ms)
        await locator.page.wait_for_timeout(settle_ms)
        return True
    except Exception:
        return False


async def _click_text_any(
    page: Page,
    texts: tuple[str, ...] | list[str],
    timeout_ms: int,
    settle_ms: int = 700,
) -> None:
    for text in texts:
        candidates = (
            page.get_by_role(
                "button",
                name=re.compile(f"^{re.escape(text)}$", re.IGNORECASE),
            ).first,
            page.locator(f'button:has-text("{text}")').first,
            page.locator(f'a:has-text("{text}")').first,
            page.locator(f'span.ui-button-text:has-text("{text}")').first,
            page.get_by_text(text, exact=True).first,
        )

        for candidate in candidates:
            if await _click_locator(candidate, timeout_ms=1800, settle_ms=settle_ms):
                await _wait_ajax(page, timeout_ms)
                return

    raise RuntimeError(f"Não encontrei nenhum botão/texto entre: {texts}")


async def _click_search_button(page: Page, options: BotOptions) -> None:
    texts = ("Buscar", "Pesquisar")

    candidates: list[Locator] = []

    for text in texts:
        candidates.extend(
            [
                page.get_by_role(
                    "button",
                    name=re.compile(f"^{re.escape(text)}$", re.IGNORECASE),
                ).first,
                page.locator(f'button:has(span.ui-button-text:has-text("{text}"))').first,
                page.locator(f'button:has-text("{text}")').first,
                page.locator(f'a:has(span.ui-button-text:has-text("{text}"))').first,
                page.locator(f'a:has-text("{text}")').first,
                page.locator(
                    f'xpath=//span[contains(@class,"ui-button-text") '
                    f'and normalize-space()="{text}"]/ancestor::button[1]'
                ).first,
                page.locator(
                    f'xpath=//span[contains(@class,"ui-button-text") '
                    f'and normalize-space()="{text}"]/ancestor::a[1]'
                ).first,
                page.locator(f'input[type="submit"][value="{text}"]').first,
            ]
        )

    for candidate in candidates:
        if await _click_locator(
            candidate,
            timeout_ms=2200,
            settle_ms=options.settle_ms,
        ):
            await _wait_ajax(page, options.timeout_ms)
            await _settle(page, options)
            return

    try:
        await page.screenshot(path="erro_botao_buscar.png", full_page=True)
    except Exception:
        pass

    raise RuntimeError(
        "Não encontrei o botão Buscar/Pesquisar. "
        "Salvei um print em erro_botao_buscar.png."
    )


async def _get_invoice_filter(page: Page, timeout_ms: int) -> Locator:
    last_error: Exception | None = None

    for selector in INVOICE_FILTER_SELECTORS:
        locator = page.locator(selector).first

        try:
            await locator.wait_for(state="visible", timeout=timeout_ms)
            return locator
        except Exception as exc:
            last_error = exc

    try:
        await page.screenshot(path="erro_filtro_numero.png", full_page=True)
    except Exception:
        pass

    raise RuntimeError(
        "Não encontrei o campo de filtro da coluna Número. "
        "Salvei um print em erro_filtro_numero.png."
    ) from last_error


async def _wait_faturas_page_ready(page: Page, timeout_ms: int) -> None:
    await _wait_ajax(page, timeout_ms)

    try:
        await _get_invoice_filter(page, timeout_ms=8000)
        return
    except Exception:
        pass

    try:
        await page.locator(SELECTORS.table_body).first.wait_for(
            state="attached",
            timeout=8000,
        )
        return
    except Exception:
        pass

    try:
        await page.screenshot(path="erro_tela_faturas.png", full_page=True)
    except Exception:
        pass

    raise RuntimeError(
        "Não encontrei a tela de faturas. "
        f"URL atual: {page.url}. "
        "Verifique o print erro_tela_faturas.png."
    )


async def _login(page: Page, options: BotOptions, log: LogFn | None = None) -> None:
    await _safe_log(log, "Abrindo tela de login...")

    await page.goto(
        PORTAL.login_url,
        wait_until="domcontentloaded",
        timeout=options.timeout_ms,
    )

    await page.locator(SELECTORS.login_input).wait_for(
        state="visible",
        timeout=options.timeout_ms,
    )

    await page.locator(SELECTORS.login_input).fill(options.username)
    await page.locator(SELECTORS.password_input).fill(options.password)

    await _settle(page, options, 400)

    await _click_text_any(
        page,
        ("Entrar",),
        options.timeout_ms,
        settle_ms=options.settle_ms,
    )

    await _wait_ajax(page, options.timeout_ms)
    await _settle(page, options)

    if await _is_login_visible(page):
        raise RuntimeError("Login não foi concluído. Verifique usuário e senha.")

    await page.goto(
        PORTAL.list_url,
        wait_until="domcontentloaded",
        timeout=options.timeout_ms,
    )

    await _wait_faturas_page_ready(page, options.timeout_ms)
    await _settle(page, options)

    await _safe_log(log, "Login realizado e tela de faturas carregada.")


async def _open_filters_if_needed(page: Page, options: BotOptions) -> None:
    try:
        if await page.locator(SELECTORS.date_from).is_visible(timeout=2000):
            return
    except Exception:
        pass

    try:
        await _click_text_any(
            page,
            ("Filtros",),
            options.timeout_ms,
            settle_ms=options.settle_ms,
        )

        await page.locator(SELECTORS.date_from).wait_for(
            state="visible",
            timeout=options.timeout_ms,
        )
        await _settle(page, options)
        return
    except Exception:
        pass

    await page.evaluate(
        """
        () => {
            const content = document.querySelector('[id="form:filtro_content"]');
            const collapsed = document.querySelector('[id="form:filtro_collapsed"]');

            if (content) {
                content.style.display = 'block';
            }

            if (collapsed) {
                collapsed.value = 'false';
            }
        }
        """
    )

    await page.locator(SELECTORS.date_from).wait_for(
        state="attached",
        timeout=options.timeout_ms,
    )

    await _settle(page, options)


async def _fill_or_set_value(page: Page, selector: str, value: str, options: BotOptions) -> None:
    locator = page.locator(selector).first

    try:
        await locator.wait_for(state="visible", timeout=3000)
        await locator.fill(value)
        await _settle(page, options, 250)
        return
    except Exception:
        pass

    await page.evaluate(
        """
        ([selector, value]) => {
            const el = document.querySelector(selector);

            if (!el) {
                throw new Error(`Elemento não encontrado: ${selector}`);
            }

            el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
        }
        """,
        [selector, value],
    )

    await _settle(page, options, 350)


async def _ensure_checkbox_checked_by_label(
    page: Page,
    label_text: str,
    options: BotOptions,
) -> bool:
    labels = page.locator(f'label:has-text("{label_text}")')
    count = await labels.count()

    if count == 0:
        return False

    label = labels.first
    for_id = await label.get_attribute("for")

    if not for_id:
        return False

    input_locator = page.locator(f'input[id="{for_id}"]').first
    await input_locator.wait_for(state="attached", timeout=options.timeout_ms)

    if await input_locator.is_checked():
        return True

    visible_box = page.locator(
        f'xpath=//input[@id="{for_id}"]'
        "/ancestor::div[contains(@class,'ui-chkbox')]"
        "//div[contains(@class,'ui-chkbox-box')]"
    ).first

    await visible_box.click(timeout=options.timeout_ms)
    await _wait_ajax(page, options.timeout_ms)
    await _settle(page, options, 300)

    return True


async def _ensure_all_situacao_checked(page: Page, options: BotOptions) -> None:
    for label in STATUS_LABELS:
        try:
            await _ensure_checkbox_checked_by_label(page, label, options)
        except Exception:
            pass

    for idx in range(6):
        input_selector = f'[id="form:j_idt596:{idx}"]'
        checkbox = page.locator(input_selector).first

        try:
            await checkbox.wait_for(state="attached", timeout=700)
        except Exception:
            continue

        try:
            if await checkbox.is_checked():
                continue

            visible_box = page.locator(
                f"xpath=//input[@id='form:j_idt596:{idx}']"
                "/ancestor::div[contains(@class,'ui-chkbox')]"
                "//div[contains(@class,'ui-chkbox-box')]"
            ).first

            await visible_box.click(timeout=options.timeout_ms)
            await _wait_ajax(page, options.timeout_ms)
            await _settle(page, options, 300)
        except Exception:
            continue


async def _prepare_list_page(
    page: Page,
    dt_from: datetime,
    options: BotOptions,
    log: LogFn | None = None,
) -> None:
    await page.goto(
        PORTAL.list_url,
        wait_until="domcontentloaded",
        timeout=options.timeout_ms,
    )

    await _wait_ajax(page, options.timeout_ms)
    await _settle(page, options)

    if await _is_login_visible(page):
        await _login(page, options, log=None)

    await _wait_faturas_page_ready(page, options.timeout_ms)
    await _open_filters_if_needed(page, options)

    await _fill_or_set_value(
        page,
        SELECTORS.date_from,
        _format_from(dt_from),
        options,
    )

    # Regra atual:
    # NÃO altera a data final.
    # O valor que estiver no portal será mantido exatamente como está.
    await _settle(page, options, 300)

    await _ensure_all_situacao_checked(page, options)

    try:
        await page.locator(ROWS_PER_PAGE_SELECTOR).first.select_option(
            "500",
            timeout=3000,
        )
        await _wait_ajax(page, options.timeout_ms)
        await _settle(page, options, 500)
    except Exception:
        pass

    await _click_search_button(page, options)

    await page.locator(SELECTORS.table_body).first.wait_for(
        state="attached",
        timeout=options.timeout_ms,
    )

    await _get_invoice_filter(page, options.timeout_ms)
    await _settle(page, options)

    await _safe_log(
        log,
        f"Aba preparada com data inicial {_format_from(dt_from)} e data final mantida como estava no portal.",
    )


async def _trigger_table_filter(page: Page, options: BotOptions) -> None:
    try:
        invoice_filter = await _get_invoice_filter(page, options.timeout_ms)
        await invoice_filter.press("Enter")
    except Exception:
        await page.keyboard.press("Enter")

    try:
        await page.evaluate(
            """
            () => {
                if (window.PF && PF('tablelist')) {
                    PF('tablelist').filter();
                }
            }
            """
        )
    except Exception:
        pass

    await _wait_ajax(page, options.timeout_ms)
    await _settle(page, options)


async def _read_reason_and_reset(
    page: Page,
    row: Locator,
    dt_from: datetime,
    options: BotOptions,
) -> str:
    link = row.locator("td").nth(0).locator("a").first

    await link.wait_for(state="visible", timeout=options.timeout_ms)
    await link.click(timeout=options.timeout_ms)

    await _wait_ajax(page, options.timeout_ms)
    await _settle(page, options)

    reason = ""

    try:
        textarea = page.locator(SELECTORS.reason_textarea).first
        await textarea.wait_for(state="attached", timeout=options.timeout_ms)

        try:
            reason = (await textarea.input_value()).strip()
        except Exception:
            reason = (await textarea.inner_text()).strip()

    except Exception:
        reason = ""

    await _prepare_list_page(page, dt_from, options, log=None)

    return reason


async def _consult_invoice(
    page: Page,
    task: InvoiceTask,
    dt_from: datetime,
    options: BotOptions,
) -> InvoiceResult:
    invoice = str(task.invoice).strip()

    if not invoice or invoice == "-":
        return InvoiceResult(status=EXCEL.value_not_found)

    invoice_filter = await _get_invoice_filter(page, options.timeout_ms)

    await invoice_filter.fill("")
    await _settle(page, options, 200)

    await invoice_filter.fill(invoice)
    await _settle(page, options, 250)

    await _trigger_table_filter(page, options)

    try:
        await page.locator(SELECTORS.table_body).first.wait_for(
            state="attached",
            timeout=options.timeout_ms,
        )
    except PlaywrightTimeoutError:
        return InvoiceResult(status=EXCEL.value_not_found)

    rows = page.locator(SELECTORS.table_rows)
    count = await rows.count()

    if count == 0:
        return InvoiceResult(status=EXCEL.value_not_found)

    for idx in range(count):
        row = rows.nth(idx)
        text = (await row.inner_text()).strip()

        if (
            "Nenhum registro" in text
            or "Nenhum dado" in text
            or "No records" in text
            or "No data" in text
        ):
            return InvoiceResult(status=EXCEL.value_not_found)

        cells = row.locator("td")
        cell_count = await cells.count()

        if cell_count < 2:
            continue

        number_text = (await cells.nth(0).inner_text()).strip()

        if number_text != invoice:
            continue

        status_text = (await cells.nth(cell_count - 1).inner_text()).strip()

        if not status_text:
            return InvoiceResult(status=EXCEL.value_not_found)

        if status_text.casefold() == "rejeitada":
            motivo = await _read_reason_and_reset(
                page=page,
                row=row,
                dt_from=dt_from,
                options=options,
            )

            return InvoiceResult(status=status_text, motivo=motivo)

        return InvoiceResult(status=status_text)

    return InvoiceResult(status=EXCEL.value_not_found)


def _split_evenly(items: list[InvoiceTask], parts: int) -> list[list[InvoiceTask]]:
    parts = max(1, min(parts, len(items)))
    chunks = [[] for _ in range(parts)]

    for idx, item in enumerate(items):
        chunks[idx % parts].append(item)

    return chunks


async def _worker(
    worker_id: int,
    context: BrowserContext,
    tasks: list[InvoiceTask],
    dt_from: datetime,
    options: BotOptions,
    results: dict[str, InvoiceResult],
    results_lock: asyncio.Lock,
    progress_state: dict[str, int],
    progress_lock: asyncio.Lock,
    total: int,
    log: LogFn | None,
    progress: ProgressFn | None,
) -> None:
    await asyncio.sleep((worker_id - 1) * 0.8)

    page = await context.new_page()

    try:
        await _prepare_list_page(page, dt_from, options, log=None)

        await _safe_log(
            log,
            f"Aba {worker_id}: iniciando {len(tasks)} consulta(s).",
        )

        for task in tasks:
            try:
                result = await _consult_invoice(page, task, dt_from, options)
            except Exception as exc:
                result = InvoiceResult(status="Erro na consulta")

                try:
                    await page.screenshot(
                        path=f"erro_fatura_{task.invoice}_aba_{worker_id}.png",
                        full_page=True,
                    )
                except Exception:
                    pass

                await _safe_log(
                    log,
                    f"Aba {worker_id}: erro na fatura {task.invoice}: {exc}",
                )

                try:
                    await _prepare_list_page(page, dt_from, options, log=None)
                except Exception:
                    pass

            async with results_lock:
                results[task.invoice] = result

            async with progress_lock:
                progress_state["done"] += 1
                done = progress_state["done"]

            if progress:
                progress(done, total)

            if result.motivo:
                await _safe_log(
                    log,
                    f"Aba {worker_id}: {task.invoice} -> {result.status} | motivo capturado.",
                )
            else:
                await _safe_log(
                    log,
                    f"Aba {worker_id}: {task.invoice} -> {result.status}",
                )

            await _settle(page, options, 500)

    finally:
        await page.close()


async def _route_fast_mode(route: Route) -> None:
    resource_type = route.request.resource_type

    if resource_type in {"image", "media", "font"}:
        await route.abort()
        return

    await route.continue_()


async def run_consultas(
    tasks: list[InvoiceTask],
    dt_from: datetime,
    options: BotOptions,
    log: LogFn | None = None,
    progress: ProgressFn | None = None,
) -> dict[str, InvoiceResult]:
    if not tasks:
        return {}

    max_tabs = max(1, min(int(options.max_tabs), 5, len(tasks)))

    results: dict[str, InvoiceResult] = {}
    results_lock = asyncio.Lock()
    progress_lock = asyncio.Lock()
    progress_state = {"done": 0}

    async with async_playwright() as p:
        try:
            browser: Browser = await p.chromium.launch(
                headless=options.headless,
                slow_mo=options.slow_mo_ms,
            )
        except Exception as exc:
            exc_str = str(exc)
            if "Executable doesn't exist" in exc_str or "playwright install" in exc_str.lower():
                await _safe_log(log, "Navegador Chromium não encontrado no sistema.")
                await _safe_log(log, "Iniciando download automático do Chromium (isso pode demorar de 1 a 2 minutos)...")
                import sys
                import playwright.__main__
                orig_argv = sys.argv
                try:
                    sys.argv = ["playwright", "install", "chromium"]
                    playwright.__main__.main()
                    await _safe_log(log, "Download do Chromium concluído com sucesso!")
                except Exception as install_exc:
                    await _safe_log(log, f"Falha ao baixar o Chromium automaticamente: {install_exc}")
                    raise RuntimeError("Falha ao instalar o navegador do Playwright. Verifique sua conexão.") from install_exc
                finally:
                    sys.argv = orig_argv

                browser = await p.chromium.launch(
                    headless=options.headless,
                    slow_mo=options.slow_mo_ms,
                )
            else:
                raise exc

        context = await browser.new_context(
            viewport={"width": 1366, "height": 768},
        )

        context.set_default_timeout(options.timeout_ms)

        if options.fast_mode:
            await context.route("**/*", _route_fast_mode)

        try:
            page = await context.new_page()

            try:
                await _login(page, options, log=log)
            finally:
                await page.close()

            chunks = _split_evenly(tasks, max_tabs)

            await _safe_log(
                log,
                f"Rodando com {len(chunks)} aba(s) simultânea(s). "
                "Limite máximo: 5.",
            )

            await asyncio.gather(
                *(
                    _worker(
                        worker_id=i + 1,
                        context=context,
                        tasks=chunk,
                        dt_from=dt_from,
                        options=options,
                        results=results,
                        results_lock=results_lock,
                        progress_state=progress_state,
                        progress_lock=progress_lock,
                        total=len(tasks),
                        log=log,
                        progress=progress,
                    )
                    for i, chunk in enumerate(chunks)
                )
            )

        finally:
            await browser.close()

    return results