import argparse
import hashlib
import logging
import os
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import requests
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
    TimeoutException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

try:
    from webdriver_manager.chrome import ChromeDriverManager  # type: ignore
except Exception:  # pragma: no cover
    ChromeDriverManager = None  # type: ignore


AUTH_URL = "https://auth.kontur.ru/?customize=diadoc&back=https%3A%2F%2Fdiadoc.kontur.ru%2F"
EMAIL = "group4@buhcentr.com"
PASSWORD = "z*o%b5g3"
SELECTION_URL_PREFIX = "https://diadoc.kontur.ru/Box/Selection"
SUPABASE_URL = "https://jnjngpeyqfmhtrmmeizy.supabase.co/rest/v1/inn_diadok?select=*"
SUPABASE_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Impuam5n"
    "cGV5cWZtaHRybW1laXp5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjA1NjM3NjEsImV4cCI6MjA3"
    "NjEzOTc2MX0.7LP-4XIodjsltDvcRefzKh_MIXuBnwtkcESG2tO0Zas"
)
SUPABASE_DATA_URL = "https://jnjngpeyqfmhtrmmeizy.supabase.co/rest/v1/data_diadok"
HEADLESS = os.environ.get("HEADLESS", "").strip().lower() in {"1", "true", "yes"}
ACTION_DELAY = float(os.environ.get("ACTION_DELAY", "0"))
DOWNLOAD_WAIT = float(os.environ.get("DOWNLOAD_WAIT", "7"))
DOWNLOAD_DIR = Path(os.environ.get("DOWNLOAD_DIR", Path.cwd() / "downloads")).resolve()
DOM_DUMP_DIR = Path(os.environ.get("DOM_DUMP_DIR", Path.cwd())).resolve()
LOG_DIR = Path(os.environ.get("LOG_DIR", Path.cwd() / "logs")).resolve()
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE_PATH = LOG_DIR / f"kontur-{datetime.now().strftime('%Y%m%d-%H%M%S')}.log"
LOG_LEVEL_NAME = os.environ.get("LOG_LEVEL", "INFO").upper()
LOOKIN_API_URL = "https://api.lookin.team/api/rest/v1/payment_documents"
LOOKIN_API_TOKEN = "zwbs1kjlVbKSMMdhKZU7HwHsfzETSYBxtVVvYhaBxDLJW"


def _setup_logger() -> logging.Logger:
    logger = logging.getLogger("kontur")
    if logger.handlers:
        return logger
    log_level = getattr(logging, LOG_LEVEL_NAME, logging.INFO)
    logger.setLevel(log_level)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    file_handler = logging.FileHandler(LOG_FILE_PATH, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    logger.addHandler(file_handler)
    logger.propagate = False
    logger.info("Логирование запущено, файл: %s", LOG_FILE_PATH)
    return logger


LOGGER = _setup_logger()


def log_info(message: str, *args: Any) -> None:
    LOGGER.info(message, *args)


def log_debug(message: str, *args: Any) -> None:
    LOGGER.debug(message, *args)


def log_error(message: str, *args: Any) -> None:
    LOGGER.error(message, *args)


def build_driver() -> webdriver.Chrome:
    log_info("Запуск Chrome-драйвера (headless=%s).", HEADLESS)
    options = Options()
    if HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--incognito")

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("prefs", prefs)

    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin:
        options.binary_location = chrome_bin

    driver_path = os.environ.get("CHROMEDRIVER")
    if driver_path and Path(driver_path).exists():
        log_info("Используем chromedriver из переменной окружения: %s", driver_path)
        service = Service(driver_path)
    else:
        if ChromeDriverManager is None:
            raise RuntimeError("Chromedriver not found and webdriver-manager is unavailable.")
        log_info("Chromedriver не указан, скачиваем через webdriver-manager.")
        service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=options)
    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(DOWNLOAD_DIR)},
        )
    except Exception:
        pass
    log_info("Chrome-драйвер готов. Каталог загрузок: %s", DOWNLOAD_DIR)
    return driver


def pause_action(label: str) -> None:
    if ACTION_DELAY <= 0:
        return
    log_info("[PAUSE] %s", label)
    time.sleep(ACTION_DELAY)


def fetch_target_inns() -> List[str]:
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "accept": "application/json",
    }
    log_info("Запрашиваем список ИНН из Supabase (%s)", SUPABASE_URL)
    response = requests.get(SUPABASE_URL, headers=headers, timeout=30)
    response.raise_for_status()
    data = response.json()
    inns: List[str] = []
    for row in data:
        raw = row.get("inn")
        if raw is None:
            continue
        inns.append(str(raw).strip())
    log_info("Получено %d ИНН для обработки.", len(inns))
    return inns


def open_auth_page(driver: webdriver.Chrome) -> None:
    log_info("Открываем страницу авторизации %s", AUTH_URL)
    driver.get(AUTH_URL)
    wait = WebDriverWait(driver, 30)

    password_tab = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//a[@role='tab' and @data-tid='tab_login']"
                "[.//span[contains(translate(normalize-space(.),'ПАРОЛЬ','пароль'),'пароль')]]",
            )
        )
    )
    password_tab.click()

    email_input = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//input[@type='email' or contains(translate(@name,'EMAIL','email'),'email') or contains(translate(@id,'EMAIL','email'),'email')]",
            )
        )
    )
    email_input.clear()
    email_input.send_keys(Keys.CONTROL, "a")
    email_input.send_keys(Keys.BACKSPACE)
    email_input.send_keys(EMAIL)

    password_input = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='password' and not(@disabled)]"))
    )
    password_input.clear()
    password_input.send_keys(Keys.CONTROL, "a")
    password_input.send_keys(Keys.BACKSPACE)
    password_input.send_keys(PASSWORD)

    login_button = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//button[contains(translate(normalize-space(.),'ВОЙТИ','войти'),'войти')]",
            )
        )
    )
    log_info("Выполняем вход в аккаунт.")
    login_button.click()

    time.sleep(5)
    log_info("Авторизация завершена, ожидаем страницу выбора ящика.")


def wait_for_selection_page(driver: webdriver.Chrome) -> None:
    wait = WebDriverWait(driver, 60)
    wait.until(lambda d: SELECTION_URL_PREFIX in d.current_url)
    current = driver.current_url
    if not current.startswith(SELECTION_URL_PREFIX):
        raise RuntimeError(f"Ожидалась страница выбора ящика, но получен URL: {current}")
    log_info("Страница выбора ящика загружена: %s", current)


def process_companies(
    driver: webdriver.Chrome, inns: List[str], date_mode: bool = False, use_lookin: bool = False
) -> None:
    if not inns:
        log_info("Список ИНН пуст, нечего открывать.")
        return
    log_info("Начинаем обработку %d ИНН.", len(inns))
    for inn in inns:
        inn_str = str(inn).strip()
        if not inn_str:
            continue
        log_info("Открываем карточку ИНН %s", inn_str)
        try:
            opened = _open_company_card(driver, inn_str, date_mode=date_mode, use_lookin=use_lookin)
        except Exception as exc:
            log_error("Ошибка при обработке карточки ИНН %s: %s", inn_str, exc)
            continue
        if not opened:
            log_error("Не удалось найти карточку для ИНН %s", inn_str)
            continue


def _open_company_card(
    driver: webdriver.Chrome, inn_str: str, date_mode: bool = False, use_lookin: bool = False
) -> bool:
    wait = WebDriverWait(driver, 30)
    log_debug("Ищем карточку по ИНН %s", inn_str)
    locator = (
        By.XPATH,
        (
            f"(//a[.//text()[contains(.,'{inn_str}')]]"
            f"|//button[.//text()[contains(.,'{inn_str}')]]"
            f"|//div[@role='button'][.//text()[contains(.,'{inn_str}')]])[1]"
        ),
    )
    try:
        element = wait.until(EC.element_to_be_clickable(locator))
    except TimeoutException:
        log_error("Карточка с ИНН %s не найдена на странице выбора.", inn_str)
        return False

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    element.click()
    try:
        wait.until(lambda d: "Box/Selection" not in d.current_url)
    except TimeoutException:
        pass
    time.sleep(2)
    log_info("Карточка ИНН %s открыта, начинаем обработку документов.", inn_str)
    success = True
    try:
        handle_documents_for_inn(driver, inn_str, date_mode=date_mode, use_lookin=use_lookin)
    except Exception as exc:
        success = False
        log_error(
            "Во время обработки документов ИНН %s произошла ошибка: %r",
            inn_str,
            exc,
        )
    finally:
        time.sleep(DOWNLOAD_WAIT)
        if not _return_to_selection_page(driver, inn_str):
            return False
    return success


def _return_to_selection_page(driver: webdriver.Chrome, inn_str: str) -> bool:
    for attempt in range(1, 4):
        try:
            if attempt > 1:
                log_debug(
                    "Повторная попытка возврата к списку ящиков после ИНН %s (попытка %d).",
                    inn_str,
                    attempt,
                )
            driver.back()
            wait_for_selection_page(driver)
            log_info("Вернулись к списку ящиков после ИНН %s.", inn_str)
            return True
        except Exception as exc:
            log_error(
                "Не удалось вернуться к списку ящиков после ИНН %s (попытка %d): %r",
                inn_str,
                attempt,
                exc,
            )
            time.sleep(3)
    try:
        log_info("Пробуем принудительно открыть страницу выбора ящиков.")
        driver.get(SELECTION_URL_PREFIX)
        wait_for_selection_page(driver)
        return True
    except Exception as exc:
        log_error(
            "Не смогли открыть страницу выбора ящика после ИНН %s: %r",
            inn_str,
            exc,
        )
        return False


def handle_documents_for_inn(
    driver: webdriver.Chrome, inn_str: str, date_mode: bool = False, use_lookin: bool = False
) -> None:
    log_info("Начинаем обработку документов для ИНН %s", inn_str)
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//body"))
        )
    except TimeoutException:
        log_error("Страница ИНН %s не загрузилась.", inn_str)
        return
    try:
        if not date_mode:
            log_info("Применяем фильтр 'за месяц' для ИНН %s", inn_str)
            apply_month_period_filter(driver)
        else:
            log_info("Режим обновления дат: фильтр не применяется для ИНН %s", inn_str)
        documents = _collect_documents_from_page(driver, inn_str, date_mode=date_mode)
    except Exception as exc:
        log_error("Не удалось применить фильтр по периоду для ИНН %s: %s", inn_str, exc)
        documents = []
    if not documents:
        log_info("Документы для ИНН %s не найдены после фильтрации.", inn_str)
        return
    try:
        _sync_documents_with_supabase(documents, date_mode=date_mode, use_lookin=use_lookin)
    except Exception as exc:
        log_error("Не удалось отправить документы в Supabase для ИНН %s: %s", inn_str, exc)
    # has_docs = select_yesterday_documents(driver)
    # if not has_docs:
    #     print(f"Документы со статусом 'вчера' не найдены для ИНН {inn_str}")
    #     return
    # try:
    #     download_btn = _find_download_button(driver)
    #     highlight_element(driver, download_btn, "rgba(0, 128, 0, 0.5)")
    #     driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", download_btn)
    #     download_btn.click()
    #
    #     format_btn = _find_original_format_button(driver)
    #     highlight_element(driver, format_btn, "rgba(0, 128, 0, 0.5)")
    #     format_btn.click()
    #     time.sleep(DOWNLOAD_WAIT)
    # except TimeoutException:
    #     print(f"Кнопка скачивания недоступна для ИНН {inn_str}")


def apply_month_period_filter(driver: webdriver.Chrome) -> None:
    log_debug("Открываем диалог фильтров и выбираем период 'Месяц'.")
    filtered_dialog = _open_filters_dialog(driver)
    _select_period_radiobutton(driver, filtered_dialog)
    _open_period_selector(driver, filtered_dialog)
    _choose_month_option(driver)
    _apply_filters(driver, filtered_dialog)


def _dialog_xpath() -> str:
    return (
        "//div[@role='dialog' or contains(@data-tid,'Modal__content') or "
        "contains(@data-tid,'ModalContainer') or contains(@class,'Modal')]"
    )


def _dump_current_dom(driver: webdriver.Chrome, inn_str: str) -> None:
    DOM_DUMP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"dom-{inn_str}-{timestamp}.html"
    output_path = DOM_DUMP_DIR / filename
    try:
        html = driver.execute_script("return document.documentElement.outerHTML;")
        output_path.write_text(html, encoding="utf-8")
        log_info("DOM текущей страницы сохранён в %s", output_path)
    except Exception as exc:
        log_error("Не удалось сохранить DOM страницы для %s: %s", inn_str, exc)


def _period_label_xpath() -> str:
    text_condition = "contains(translate(normalize-space(.),'ПЕРИОД','период'),'период')"
    return (
        f"{_dialog_xpath()}//label[{text_condition}]|"
        f"{_dialog_xpath()}//span[{text_condition}]/ancestor::label[1]"
    )


def _period_selector_xpath() -> str:
    selector_condition = (
        "self::select or @role='combobox' or "
        "(@role='button' and (@aria-haspopup='listbox' or contains(@data-tid,'Select') "
        "or contains(@class,'Select'))) or contains(@data-tid,'Select__control') "
        "or contains(@class,'Select__control')"
    )
    return f"({_period_label_xpath()})/following::*[{selector_condition}][1]"


def _generic_selector_xpath() -> str:
    return (
        f"{_dialog_xpath()}//select | "
        f"{_dialog_xpath()}//*[@role='combobox'] | "
        f"{_dialog_xpath()}//button[contains(@data-tid,'Select') or @aria-haspopup='listbox' "
        "or contains(@class,'Select')] | "
        f"{_dialog_xpath()}//div[contains(@data-tid,'Select__control') or contains(@class,'Select')]"
        "[@role='button' or @role='combobox']"
    )


def _not_selected_selector_xpath() -> str:
    text_cond = _text_condition("не выбрано")
    scoped = (
        f"({_period_label_xpath()})/following::*[{text_cond}]"
        "[self::span or self::div or self::button or self::label or self::p or self::option]"
    )
    clickable = (
        f"{scoped}/ancestor-or-self::button[1]|"
        f"{scoped}/ancestor-or-self::*[@role='button' or @role='combobox'][1]|"
        f"{scoped}/ancestor-or-self::select[1]|"
        f"{scoped}/ancestor-or-self::label[1]|"
        f"{scoped}/ancestor-or-self::div[contains(@class,'Select') or contains(@data-tid,'Select')][1]"
    )
    return f"{clickable}|{scoped}"


def _text_condition(term: str) -> str:
    lower = term.lower()
    upper = term.upper()
    return f"contains(translate(normalize-space(.),'{upper}','{lower}'),'{lower}')"


def _month_text_condition() -> str:
    variants = [
        "за месяц",
        "текущий месяц",
        "этот месяц",
        "месяц",
    ]
    return " or ".join(_text_condition(term) for term in variants)


def _month_option_xpath() -> str:
    base = (
        "//div[@role='option' or contains(@data-tid,'MenuItem') or contains(@class,'MenuItem') "
        "or contains(@class,'Select__option') or contains(@class,'Option')]"
        "|//li[@role='option' or contains(@data-tid,'MenuItem') or contains(@class,'MenuItem')]"
        "|//button[@role='option' or contains(@data-tid,'MenuItem') or contains(@class,'MenuItem') "
        "or contains(@class,'Select__option')]"
        "|//span[@role='option' or ancestor::li[@role='option'] or ancestor::div[@role='option']]"
        "|//option"
    )
    return f"({base})[{_month_text_condition()}]"


def _open_filters_dialog(driver: webdriver.Chrome) -> WebElement:
    wait = WebDriverWait(driver, 20)
    log_info("Открываем диалог фильтров документов.")
    filters_btn = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//button[contains(translate(normalize-space(.),'ФИЛЬТР','фильтр'),'фильтр')]",
            )
        )
    )
    highlight_element(driver, filters_btn, "rgba(0, 128, 0, 0.25)")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", filters_btn)
    filters_btn.click()
    dialog = wait.until(EC.presence_of_element_located((By.XPATH, _dialog_xpath())))
    return dialog


def _select_period_radiobutton(driver: webdriver.Chrome, dialog: WebElement) -> None:
    log_debug("Выбираем радиокнопку 'Период'.")
    wait = WebDriverWait(driver, 15)
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dialog)
    period_radio = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                _period_label_xpath(),
            )
        )
    )
    highlight_element(driver, period_radio, "rgba(0, 128, 0, 0.25)")
    period_radio.click()
    pause_action("Выбор радиокнопки 'Период'")


def _open_period_selector(driver: webdriver.Chrome, dialog: WebElement) -> None:
    log_debug("Открываем выпадающий список периодов.")
    wait = WebDriverWait(driver, 15)
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dialog)
    selector = None
    locator_chain = [
        _not_selected_selector_xpath(),
        _period_selector_xpath(),
        _generic_selector_xpath(),
    ]
    last_exc: Optional[Exception] = None
    for locator in locator_chain:
        try:
            selector = wait.until(EC.element_to_be_clickable((By.XPATH, locator)))
            break
        except TimeoutException as exc:
            last_exc = exc
    if selector is None:
        if last_exc:
            raise last_exc
        raise TimeoutException("Не найден элемент для открытия списка периода")
    highlight_element(driver, selector, "rgba(0, 128, 0, 0.25)")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", selector)
    try:
        selector.click()
    except Exception:
        driver.execute_script("arguments[0].click();", selector)
    pause_action("Открытие списка периода")


def _choose_month_option(driver: webdriver.Chrome) -> None:
    log_debug("Выбираем опцию периода 'За месяц'.")
    wait = WebDriverWait(driver, 15)
    try:
        option = wait.until(EC.element_to_be_clickable((By.XPATH, _month_option_xpath())))
    except TimeoutException:
        options = driver.find_elements(By.XPATH, _month_option_xpath())
        option = options[0] if options else None
        if option is None:
            raise
    highlight_element(driver, option, "rgba(0, 128, 0, 0.25)")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option)
    try:
        option.click()
    except (ElementClickInterceptedException, Exception):
        driver.execute_script("arguments[0].click();", option)
    pause_action("Выбор пункта 'за месяц'")


def _apply_filters(driver: webdriver.Chrome, dialog: WebElement) -> None:
    log_debug("Применяем выбранные фильтры.")
    wait = WebDriverWait(driver, 15)
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dialog)
    apply_btn = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                f"{_dialog_xpath()}//button[contains(translate(normalize-space(.),'ПРИМЕНИТЬ','применить'),'применить')]",
            )
        )
    )
    highlight_element(driver, apply_btn, "rgba(0, 128, 0, 0.25)")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", apply_btn)
    apply_btn.click()
    pause_action("Клик по кнопке 'Применить'")
    try:
        wait.until(EC.staleness_of(dialog))
    except Exception:
        pass


DOCUMENT_ROW_XPATH = (
    "//article[contains(@data-tid,'singleLetter') or contains(@data-tid,'letter')]|"
    "//*[@documentid or @data-document-id or @data-doc-id or @data-entity-id]"
)
UUID_RE = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")
MONTHS_MAP = {
    "января": 1,
    "февраля": 2,
    "марта": 3,
    "апреля": 4,
    "мая": 5,
    "июня": 6,
    "июля": 7,
    "августа": 8,
    "сентября": 9,
    "октября": 10,
    "ноября": 11,
    "декабря": 12,
    "январь": 1,
    "февраль": 2,
    "март": 3,
    "апрель": 4,
    "май": 5,
    "июнь": 6,
    "июль": 7,
    "август": 8,
    "сентябрь": 9,
    "октябрь": 10,
    "ноябрь": 11,
    "декабрь": 12,
    "january": 1,
    "jan": 1,
    "february": 2,
    "feb": 2,
    "march": 3,
    "mar": 3,
    "april": 4,
    "apr": 4,
    "may": 5,
    "june": 6,
    "jun": 6,
    "july": 7,
    "jul": 7,
    "august": 8,
    "aug": 8,
    "september": 9,
    "sep": 9,
    "sept": 9,
    "october": 10,
    "oct": 10,
    "november": 11,
    "nov": 11,
    "december": 12,
    "dec": 12,
}


def _normalize_ws(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def _extract_text(element: WebElement, selectors: List[str]) -> str:
    for selector in selectors:
        try:
            target = element.find_element(By.XPATH, selector)
        except NoSuchElementException:
            continue
        text = (target.text or "").strip()
        if not text:
            text = (target.get_attribute("textContent") or "").strip()
        if text:
            return _normalize_ws(text)
    return ""


def _extract_document_id(element: WebElement) -> Optional[str]:
    attrs = [
        "documentid",
        "data-documentid",
        "data-document-id",
        "data-doc-id",
        "data-entity-id",
        "data-letter-id",
        "data-letter-document-id",
        "data-letter-entry-id",
        "data-entry-id",
        "data-id",
    ]

    def _from_text(value: Optional[str], allow_generic: bool = True) -> Optional[str]:
        if not value:
            return None
        match = UUID_RE.search(value)
        if match:
            return match.group(0)
        if not allow_generic:
            return None
        cleaned = value.strip().strip("'\"")
        cleaned = cleaned.replace("\xa0", " ")
        cleaned = re.sub(r"\s+", "", cleaned)
        if not cleaned or len(cleaned) > 256:
            return None
        return cleaned
        return None

    for attr in attrs:
        doc_id = _from_text(element.get_attribute(attr))
        if doc_id:
            return doc_id

    candidate_paths = [
        ".//*[@documentid or @data-document-id or @data-doc-id or @data-entity-id or "
        "@data-letter-id or @data-letter-document-id][1]",
        ".//a[contains(@href,'documentId') or contains(@href,'documentid')][1]",
        ".//div[contains(@data-tid,'DocumentTitle') or contains(@data-tid,'DocumentName')]//a[1]",
    ]
    for path in candidate_paths:
        try:
            node = element.find_element(By.XPATH, path)
        except NoSuchElementException:
            continue
        for attr in attrs:
            doc_id = _from_text(node.get_attribute(attr))
            if doc_id:
                return doc_id
        href_doc = _from_text(node.get_attribute("href"))
        if href_doc:
            return href_doc
        text_doc = _from_text(node.get_attribute("textContent") or node.text)
        if text_doc:
            return text_doc

    outer_html = element.get_attribute("outerHTML")
    return _from_text(outer_html, allow_generic=False)


def _make_synthetic_id(
    inn_str: str,
    sender: str,
    name: str,
    date_value: Optional[str],
    row_index: int,
    node_index: int,
) -> str:
    payload = f"{inn_str}|{sender}|{name}|{date_value or ''}|{row_index}|{node_index}"
    digest = hashlib.sha1(payload.encode("utf-8")).hexdigest()
    return f"synthetic-{digest}"


def _parse_money(value: str) -> Optional[float]:
    if not value:
        return None
    cleaned = value.replace("\xa0", "").replace(" ", "")
    match = re.search(r"-?\d+(?:[.,]\d+)?", cleaned)
    if not match:
        return None
    number = match.group(0).replace(",", ".")
    try:
        return float(number)
    except ValueError:
        return None


def _parse_date_value(value: str) -> Optional[str]:
    if not value:
        return None
    normalized = value.replace("\xa0", " ").strip()
    lowered = normalized.lower()
    today = datetime.now().date()
    # если в ячейке указано только время (например, "13:45"), считаем, что дата - сегодня
    if re.fullmatch(r"\d{1,2}[:.]\d{2}", lowered):
        return today.isoformat()
    if lowered.startswith("сегодня"):
        return today.isoformat()
    if lowered.startswith("вчера"):
        return (today - timedelta(days=1)).isoformat()
    if lowered.startswith("позавчера"):
        return (today - timedelta(days=2)).isoformat()
    formats = ["%d.%m.%Y", "%d.%m.%y", "%d-%m-%Y", "%Y-%m-%d"]
    for fmt in formats:
        try:
            parsed = datetime.strptime(normalized, fmt)
            if parsed.year < 2000 and "%y" in fmt:
                parsed = parsed.replace(year=parsed.year + 2000)
            return parsed.date().isoformat()
        except ValueError:
            continue
    textual = re.search(
        r"(?P<day>\d{1,2})\s+?(?P<month>[A-Za-zА-Яа-яЁё]+)[,]?\s+(?P<year>\d{2,4})",
        normalized,
    )
    if textual:
        day = int(textual.group("day"))
        month_raw = textual.group("month").lower().strip(".,")
        month = MONTHS_MAP.get(month_raw)
        year = int(textual.group("year"))
        if year < 100:
            year += 2000
        if month:
            try:
                parsed = datetime(year, month, day)
                return parsed.date().isoformat()
            except ValueError:
                return None
    month_only = re.search(
        r"(?P<day>\d{1,2})\s+?(?P<month>[A-Za-zА-Яа-яЁё]+)",
        normalized,
    )
    if month_only:
        day = int(month_only.group("day"))
        month_raw = month_only.group("month").lower().strip(".,")
        month = MONTHS_MAP.get(month_raw)
        if month:
            year = today.year
            try:
                parsed = datetime(year, month, day)
            except ValueError:
                return None
            # если дата из будущего более чем на месяц, сдвигаем год назад
            if parsed.date() - today > timedelta(days=180):
                parsed = parsed.replace(year=year - 1)
            return parsed.date().isoformat()
    return None


def _status_to_bool(value: str) -> Optional[bool]:
    if not value:
        return None
    lowered = value.lower()
    if "подписан" in lowered:
        return True
    if "треб" in lowered or "ожида" in lowered or "не подпис" in lowered:
        return False
    return None


def _wait_for_documents_list(driver: webdriver.Chrome) -> List[WebElement]:
    wait = WebDriverWait(driver, 40)
    log_debug("Ожидаем появление списка документов.")
    wait.until(lambda d: len(d.find_elements(By.XPATH, DOCUMENT_ROW_XPATH)) > 0)
    time.sleep(1)
    elements = driver.find_elements(By.XPATH, DOCUMENT_ROW_XPATH)
    visible = [el for el in elements if el.is_displayed()]
    log_info("Отображается %d строк(и) документов на странице.", len(visible))
    return visible


def _extract_sender(driver: webdriver.Chrome, element: WebElement) -> Tuple[str, Optional[WebElement]]:
    selectors = [
        ".//span[@data-tid='DocumentMetadata__counteragent']",
        ".//span[@data-tid='DocumentMetadata__correspondent']",
        ".//span[@data-tid='DocumentMetadata__participant']",
        ".//span[@data-tid='DocumentMetadata__sender']",
        ".//span[@data-tid='DocumentMetadata__organization']",
        ".//span[contains(@data-tid,'CounteragentName')]",
        ".//span[contains(@data-tid,'CorrespondentName')]",
        ".//span[contains(@data-tid,'SenderName')]",
        ".//span[contains(@data-tid,'CompanyName')]",
        ".//span[contains(@data-tid,'ParticipantName')]",
        ".//*[@data-tid='CountragentBaseView']//span",
    ]
    sender = ""
    sender_element: Optional[WebElement] = None
    for selector in selectors:
        try:
            node = element.find_element(By.XPATH, selector)
        except NoSuchElementException:
            continue
        text = (node.text or "").strip()
        if not text:
            text = (node.get_attribute("textContent") or "").strip()
        if text:
            sender = _normalize_ws(text)
            sender_element = node
            break
    if not sender:
        fallback_selectors = [
            ".//div[contains(@data-tid,'CounteragentName')]",
            ".//div[contains(@data-tid,'AuthorName')]",
            ".//div[contains(@data-tid,'SenderName')]",
            ".//div[contains(@class,'CounteragentBlock')]//div[1]",
            ".//div[contains(@data-tid,'CorrespondentBlock')]//div[1]",
        ]
        for selector in fallback_selectors:
            try:
                node = element.find_element(By.XPATH, selector)
            except NoSuchElementException:
                continue
            text = (node.text or "").strip()
            if not text:
                text = (node.get_attribute("textContent") or "").strip()
            if text:
                sender = _normalize_ws(text)
                sender_element = node
                break
    if not sender:
        text = (element.text or "").lower()
        marker_idx = text.find("отправ")
        if marker_idx != -1:
            raw_text = (element.text or "")
            sender = _normalize_ws(raw_text)
        else:
            marker_idx = text.find("корреспондент")
            if marker_idx != -1:
                sender = _normalize_ws(element.text)
    if not sender:
        try:
            column = element.find_element(
                By.XPATH,
                ".//div[@role='gridcell'][1]//span[1]|.//div[@role='gridcell'][1]",
            )
            sender = _normalize_ws(column.text)
            sender_element = column
        except NoSuchElementException:
            sender = ""
            sender_element = None
    return sender, sender_element


def _collect_sender_details(driver: webdriver.Chrome, sender_element: WebElement) -> Dict[str, Optional[str]]:
    details: Dict[str, Optional[str]] = {}
    try:
        ActionChains(driver).move_to_element(sender_element).pause(0.2).perform()
    except Exception as exc:
        log_debug("Не удалось навести курсор на отправителя: %s", exc)
        return details
    tooltip_locator = (
        By.XPATH,
        "//div[@data-tid='Popup__root']//p[contains(translate(normalize-space(.),'ИНН','инн'),'инн')]",
    )
    tooltip: Optional[WebElement] = None
    try:
        tooltip = WebDriverWait(driver, 3).until(EC.visibility_of_element_located(tooltip_locator))
    except TimeoutException:
        log_debug("Всплывающая подсказка с реквизитами отправителя не появилась.")
    if tooltip is None:
        return details
    raw_text = (tooltip.text or "").strip()
    if not raw_text:
        raw_text = (tooltip.get_attribute("textContent") or "").strip()
    if not raw_text:
        log_debug("Подсказка отображается, но текст получить не удалось.")
        return details
    parsed = _parse_sender_tooltip_text(raw_text)
    details.update(parsed)
    if details.get("inn") or details.get("kpp"):
        log_info(
            "Получены реквизиты отправителя из подсказки: ИНН=%s, КПП=%s",
            details.get("inn"),
            details.get("kpp"),
        )
    else:
        log_debug("Подсказка отправителя получена, но ИНН/КПП отсутствуют. Текст: %s", raw_text)
    try:
        ActionChains(driver).move_by_offset(20, 0).perform()
    except Exception:
        pass
    return details


def _parse_sender_tooltip_text(raw_text: str) -> Dict[str, Optional[str]]:
    cleaned = raw_text.replace("\xa0", " ").strip()
    lines = [line.strip() for line in cleaned.splitlines() if line.strip()]
    details: Dict[str, Optional[str]] = {"full_name": None, "inn": None, "kpp": None}
    if lines:
        details["full_name"] = lines[0]
    inn_match = re.search(r"ИНН:\s*(\d+)", cleaned, re.IGNORECASE)
    if inn_match:
        details["inn"] = inn_match.group(1)
    kpp_match = re.search(r"КПП:\s*(\d+)", cleaned, re.IGNORECASE)
    if kpp_match:
        details["kpp"] = kpp_match.group(1)
    return details


def _is_valid_document_name(text: str) -> bool:
    if not text:
        return False
    normalized = _normalize_ws(text)
    if len(normalized) < 3:
        return False
    lower = normalized.lower()
    forbidden_markers = [
        "₽",
        " руб",
        "ндс",
        "требует",
        "требуется",
        "подпис",
        "документооборот",
        "доверенности",
        "доверенностей",
        "история документа",
    ]
    if any(marker in lower for marker in forbidden_markers):
        return False
    return True


def _extract_names(element: WebElement) -> List[str]:
    candidates_xpath = (
        ".//div[contains(@data-tid,'DocumentName')]//a|"
        ".//div[contains(@data-tid,'DocumentName')]//span|"
        ".//*[@data-tid='DocumentMetadata__title']//a|"
        ".//*[@data-tid='DocumentMetadata__title']//span|"
        ".//*[@data-tid='DocumentMetadata__documentName']//a|"
        ".//*[@data-tid='DocumentMetadata__documentName']//span|"
        ".//*[@data-tid='documentName']//div|"
        ".//*[@data-tid='documentName']//span|"
        ".//a[contains(@href,'.pdf') or contains(@href,'.xml') or contains(@href,'/document/')]"
    )
    names: List[str] = []
    seen: Set[str] = set()
    try:
        nodes = element.find_elements(By.XPATH, candidates_xpath)
    except Exception:
        nodes = []
    for node in nodes:
        text = (node.text or "").strip()
        if not text:
            text = (node.get_attribute("textContent") or "").strip()
        normalized = _normalize_ws(text)
        if normalized and normalized not in seen and _is_valid_document_name(normalized):
            names.append(normalized)
            seen.add(normalized)
    if not names:
        fallback_selectors = [
            ".//*[@data-tid='DocumentMetadata__title']",
            ".//*[@data-tid='DocumentName']",
            ".//*[@data-tid='DocumentMetadata__documentName']",
            ".//*[@data-tid='documentName']",
        ]
        fallback = _extract_text(element, fallback_selectors)
        if fallback and fallback not in seen and _is_valid_document_name(fallback):
            names.append(fallback)
            seen.add(fallback)
    return names


def _extract_summaries(element: WebElement) -> Dict[str, Optional[float]]:
    sum_text = _extract_text(
        element,
        [
            ".//*[@data-tid='DocumentMetadata__primary']",
            ".//*[@data-tid='DocumentMetadata__sum']",
            ".//span[contains(@class,'JnKP0')]",
        ],
    )
    nds_text = _extract_text(
        element,
        [
            ".//*[@data-tid='DocumentMetadata__additional']",
            ".//span[contains(@class,'L2xKE')]",
            ".//*[contains(translate(normalize-space(.),'НДС','ндс'),'ндс')]",
        ],
    )
    return {"summ": _parse_money(sum_text), "summ_nds": _parse_money(nds_text)}


def _extract_status(element: WebElement) -> bool:
    status_text = _extract_text(
        element,
        [
            ".//*[@data-tid='DocumentStatusBadge']",
            ".//*[contains(@data-tid,'StatusBadge')]",
            ".//*[contains(@data-tid,'DocumentStatus')]",
        ],
    )
    if not status_text:
        status_text = element.get_attribute("aria-label") or ""
    bool_value = _status_to_bool(status_text)
    if bool_value is None:
        bool_value = _status_to_bool(element.text or "")
    if bool_value is None:
        bool_value = False
    return bool_value


def _extract_date(element: WebElement) -> Optional[str]:
    selectors = [
        ".//time[@data-tid='Date']",
        ".//time",
        ".//*[@data-tid='DocumentMetadata__date']",
        ".//*[@data-tid='DocumentDate']",
    ]
    for selector in selectors:
        try:
            target = element.find_element(By.XPATH, selector)
        except NoSuchElementException:
            continue
        candidates = [
            target.get_attribute("datetime"),
            target.get_attribute("data-datetime"),
            target.get_attribute("title"),
            target.text,
            target.get_attribute("textContent"),
        ]
        for candidate in candidates:
            parsed = _parse_date_value(candidate)
            if parsed:
                return parsed
    return _parse_date_value(element.text or "")


def _collect_documents_from_page(
    driver: webdriver.Chrome, inn_str: str, date_mode: bool = False
) -> List[Dict[str, Any]]:
    rows = _wait_for_documents_list(driver)
    try:
        _dump_current_dom(driver, inn_str)
    except Exception:
        pass
    documents: List[Dict[str, Any]] = []
    emitted_record_keys: Set[str] = set()
    log_info("Начинаем разбор %d строк документов для ИНН %s.", len(rows), inn_str)
    for row_index, row in enumerate(rows, start=1):
        log_debug("Обрабатываем строку #%d для ИНН %s", row_index, inn_str)
        highlight_element(driver, row, "rgba(0, 255, 0, 0.15)")
        if date_mode:
            sender = ""
            sender_element = None
            sender_details: Dict[str, Optional[str]] = {}
        else:
            sender, sender_element = _extract_sender(driver, row)
            sender_details = (
                _collect_sender_details(driver, sender_element) if sender_element else {}
            )
            popup_sender = sender_details.get("full_name") if sender_details else None
            if popup_sender:
                sender = popup_sender
        row_date = _extract_date(row)
        doc_nodes = row.find_elements(
            By.XPATH,
            ".//a[@data-documentid or @data-doc-id or @documentid]",
        )
        if not doc_nodes:
            doc_nodes = [row]
        for node_index, node in enumerate(doc_nodes, start=1):
            doc_id = _extract_document_id(node) or _extract_document_id(row)
            names = _extract_names(node)
            if not names:
                names = _extract_names(row)
            if not names:
                fallback_name = doc_id if doc_id else f"Документ #{row_index}"
                names = [fallback_name]
            summaries = (
                _extract_summaries(node)
                if not date_mode
                else {"summ": None, "summ_nds": None}
            )
            if (not summaries["summ"] and not summaries["summ_nds"]) and not date_mode:
                summaries = _extract_summaries(row)
            status = (
                _extract_status(node) if node is not row else _extract_status(row)
            ) if not date_mode else False
            date_value = _extract_date(node) or row_date
            for name in names:
                dedupe_key = "|".join(
                    [
                        doc_id or "",
                        name if not date_mode else "",
                        date_value or "",
                        sender if not date_mode else "",
                    ]
                )
                if dedupe_key in emitted_record_keys:
                    continue
                record_id = doc_id
                if not record_id:
                    record_id = _make_synthetic_id(
                        inn_str,
                        sender,
                        name,
                        date_value,
                        row_index,
                        node_index,
                    )
                    log_debug(
                        "Сгенерирован синтетический идентификатор %s для документа '%s' (строка #%d, узел #%d).",
                        record_id,
                        name,
                        row_index,
                        node_index,
                    )
                log_debug(
                    "Документ найден: id=%s, name=%s, sender=%s, date=%s, summ=%s, summ_nds=%s, status=%s",
                    record_id,
                    name,
                    sender,
                    date_value,
                    summaries["summ"],
                    summaries["summ_nds"],
                    status,
                )
                record: Dict[str, Any] = {
                    "id_doc": record_id,
                    "inn": inn_str,
                    "date": date_value,
                }
                if not date_mode:
                    record.update(
                        {
                            "sender": sender,
                            "name": name,
                            "summ": summaries["summ"],
                            "summ_nds": summaries["summ_nds"],
                            "status": status,
                        }
                    )
                    if sender_details:
                        sender_inn = sender_details.get("inn")
                        sender_kpp = sender_details.get("kpp")
                        if sender_inn:
                            record["inn_popup"] = sender_inn
                        if sender_kpp:
                            record["kpp_popup"] = sender_kpp
                documents.append(record)
                emitted_record_keys.add(dedupe_key)
    log_info("Собрано %d документов на странице для ИНН %s.", len(documents), inn_str)
    return documents


def _normalize_db_status(value: Any) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"true", "1", "yes", "да"}:
        return True
    if text in {"false", "0", "no", "нет"}:
        return False
    return None


def _fetch_existing_supabase_records(
    record_ids: List[str], select: str = "id_doc,status"
) -> Dict[str, Dict[str, Any]]:
    existing: Dict[str, Dict[str, Any]] = {}
    unique_ids = sorted({rid for rid in record_ids if rid})
    if not unique_ids:
        return existing
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "accept": "application/json",
    }
    chunk_size = 50
    for start in range(0, len(unique_ids), chunk_size):
        chunk = unique_ids[start : start + chunk_size]
        quoted_ids = ",".join(f'"{rid}"' for rid in chunk)
        params = {"select": select, "id_doc": f"in.({quoted_ids})"}
        response = requests.get(
            SUPABASE_DATA_URL,
            headers=headers,
            params=params,
            timeout=30,
        )
        response.raise_for_status()
        rows = response.json()
        for row in rows:
            rid = str(row.get("id_doc") or "").strip()
            if rid:
                existing[rid] = row
    log_info("Получено %d существующих записей из Supabase для проверки.", len(existing))
    return existing


def _insert_new_records(records: List[Dict[str, Any]]) -> None:
    if not records:
        return
    all_keys: Set[str] = set()
    for record in records:
        all_keys.update(record.keys())
    ordered_keys = sorted(all_keys)
    normalized_records = [{key: record.get(key) for key in ordered_keys} for record in records]
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
    }
    response = requests.post(
        SUPABASE_DATA_URL,
        headers=headers,
        json=normalized_records,
        timeout=30,
    )
    if response.status_code not in (200, 201, 204):
        raise RuntimeError(
            f"Supabase вернул ошибку {response.status_code}: {response.text}"
        )
    log_info("В Supabase добавлено %d новых записей.", len(records))


def _update_supabase_status(id_doc: str, status: bool) -> None:
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }
    params = {"id_doc": f"eq.{id_doc}"}
    payload = {"status": status}
    response = requests.patch(
        SUPABASE_DATA_URL,
        headers=headers,
        params=params,
        json=payload,
        timeout=30,
    )
    if response.status_code not in (200, 204):
        raise RuntimeError(
            f"Не удалось обновить статус документа {id_doc}: "
            f"{response.status_code} {response.text}"
        )
    log_info("Обновлён статус документа %s -> %s", id_doc, status)


def _update_supabase_date(id_doc: str, date_value: str) -> None:
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }
    params = {"id_doc": f"eq.{id_doc}"}
    payload = {"date": date_value}
    response = requests.patch(
        SUPABASE_DATA_URL,
        headers=headers,
        params=params,
        json=payload,
        timeout=30,
    )
    if response.status_code not in (200, 204):
        raise RuntimeError(
            f"Не удалось обновить дату документа {id_doc}: "
            f"{response.status_code} {response.text}"
        )
    log_info("Обновлена дата документа %s -> %s", id_doc, date_value)


def _post_documents_to_lookin(records: List[Dict[str, Any]]) -> None:
    if not records:
        log_info("Отправлять нечего: список документов пуст (LOOKIN).")
        return
    headers = {
        "Authorization": f"Bearer {LOOKIN_API_TOKEN}",
        "Content-Type": "application/json",
    }
    response = requests.post(
        LOOKIN_API_URL,
        headers=headers,
        json=records,
        timeout=30,
    )
    if response.status_code not in (200, 201, 204):
        raise RuntimeError(
            f"LOOKIN API вернул ошибку {response.status_code}: {response.text}"
        )
    log_info("В LOOKIN API отправлено %d записей.", len(records))


def _is_empty_date_value(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def _sync_documents_with_supabase(
    records: List[Dict[str, Any]], date_mode: bool = False, use_lookin: bool = False
) -> None:
    if not records:
        log_info("Отправлять нечего: список документов пуст.")
        return
    if use_lookin and not date_mode:
        _post_documents_to_lookin(records)
        return
    if use_lookin and date_mode:
        log_info("Режим LOOKIN не поддерживает обновление дат, пропускаем отправку.")
        return
    record_ids = [rec.get("id_doc", "") for rec in records if rec.get("id_doc")]
    if date_mode:
        existing = _fetch_existing_supabase_records(record_ids, select="id_doc,date")
        updated_dates = 0
        for record in records:
            record_id = record.get("id_doc")
            date_value = record.get("date")
            if not record_id or not date_value:
                continue
            existing_row = existing.get(record_id)
            if existing_row is None:
                log_debug("Документ %s отсутствует в Supabase, пропускаем обновление даты.", record_id)
                continue
            current_date = existing_row.get("date")
            if _is_empty_date_value(current_date):
                try:
                    _update_supabase_date(record_id, date_value)
                    updated_dates += 1
                except Exception as exc:
                    log_error("Не удалось обновить дату документа %s: %s", record_id, exc)
        if updated_dates:
            log_info("Обновлено дат документов: %d", updated_dates)
        else:
            log_info("Не найдено записей с пустой датой для обновления.")
        return

    existing = _fetch_existing_supabase_records(record_ids)
    new_records: List[Dict[str, Any]] = []
    updated_count = 0
    for record in records:
        record_id = record.get("id_doc")
        if not record_id:
            log_error("Пропускаем документ без id_doc: %s", record)
            continue
        existing_row = existing.get(record_id)
        if existing_row is None:
            new_records.append(record)
            continue
        existing_status = _normalize_db_status(existing_row.get("status"))
        incoming_status = bool(record.get("status"))
        if existing_status is None:
            existing_status = False
        if existing_status != incoming_status:
            try:
                _update_supabase_status(record_id, incoming_status)
                updated_count += 1
            except Exception as exc:
                log_error("Не удалось обновить статус документа %s: %s", record_id, exc)
    if new_records:
        _insert_new_records(new_records)
    else:
        log_info("Новых документов для загрузки в Supabase нет.")
    if updated_count:
        log_info("Обновлено статусов документов: %d", updated_count)


def select_yesterday_documents(driver: webdriver.Chrome) -> bool:
    time_elements = driver.find_elements(
        By.XPATH,
        "//time[contains(translate(normalize-space(.),'ВЧЕРА','вчера'),'вчера')]",
    )
    if not time_elements:
        return False

    any_selected = False
    for index, time_el in enumerate(time_elements, start=1):
        driver.execute_script(
            "arguments[0].style.outline='2px solid #00cc00';"
            "arguments[0].style.backgroundColor='rgba(0,204,0,0.25)';",
            time_el,
        )
        try:
            row = time_el.find_element(
                By.XPATH,
                "./ancestor::*[@role='row' or contains(@data-tid,'letter') or @data-tid='singleLetter'][1]",
            )
        except NoSuchElementException:
            log_error("Не смогли определить строку документа для элемента 'Вчера'")
            continue

        driver.execute_script(
            "arguments[0].style.backgroundColor='rgba(0,204,0,0.1)';"
            "arguments[0].style.outline='1px dashed #00aa00';",
            row,
        )

        checkbox_label: Optional[WebElement]
        checkbox_input: Optional[WebElement] = None

        try:
            checkbox_label = row.find_element(By.XPATH, ".//label[@data-tid='documentCheckbox']")
        except NoSuchElementException:
            checkbox_label = None

        if checkbox_label:
            try:
                checkbox_input = checkbox_label.find_element(By.XPATH, ".//input[@type='checkbox']")
            except NoSuchElementException:
                checkbox_input = None

        if checkbox_input is None:
            try:
                checkbox_input = row.find_element(By.XPATH, ".//input[@type='checkbox']")
            except NoSuchElementException:
                checkbox_input = None

        click_target: Optional[WebElement] = checkbox_label or checkbox_input
        if click_target is None:
            log_error("Чекбокс в строке 'Вчера' не найден.")
            continue

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", click_target)
        highlight_element(driver, click_target, "rgba(0, 128, 0, 0.5)")
        log_info("Пробуем поставить галочку для строки №%d", index)
        try:
            click_target.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", click_target)
        except Exception:
            driver.execute_script("arguments[0].click();", click_target)

        pause_action("Выбор документа (дата 'вчера')")
        any_selected = True

        if checkbox_input is not None:
            try:
                selected_state = checkbox_input.is_selected()
            except Exception:
                selected_state = bool(
                    driver.execute_script("return arguments[0].checked === true;", checkbox_input)
                )
            if not selected_state:
                driver.execute_script(
                    "arguments[0].checked = true;"
                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                    checkbox_input,
                )
    return any_selected


def _find_download_button(driver: webdriver.Chrome) -> WebElement:
    wait = WebDriverWait(driver, 20)
    xpaths = [
        "//*[@data-tid='batchDownloadDropdownButton']//button[contains(@data-tid,'Button__rootElement')]",
        "//button[contains(@data-tid,'Button__rootElement')][.//span[contains(translate(normalize-space(.),'СКАЧ','скач'),'скач')]]",
        "//button[.//*[contains(translate(normalize-space(.),'СКАЧ','скач'),'скач')]]",
        "//button[contains(translate(normalize-space(string(.)),'СКАЧ','скач'),'скач')]",
        "//span[contains(translate(normalize-space(.),'СКАЧ','скач'),'скач')]/ancestor::button[1]",
        "//button[contains(@aria-controls,'Select__menu')][.//span[contains(translate(normalize-space(.),'СКАЧ','скач'),'скач')]]",
    ]
    last_exc: Optional[Exception] = None
    for xp in xpaths:
        try:
            element = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            if element.is_displayed():
                wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
                return element
        except Exception as exc:
            last_exc = exc

    # fallback: перебираем все кнопки и ищем текст "скач"
    buttons = driver.find_elements(By.TAG_NAME, "button")
    for btn in buttons:
        try:
            text = (btn.text or "").strip().lower()
            if "скач" in text:
                return btn
        except Exception:
            continue
    js_target = driver.execute_script(
        """
        const matches = Array.from(document.querySelectorAll('[data-tid="batchDownloadDropdownButton"] button, button'));
        return matches.find(btn => {
            const text = (btn.textContent || '').toLowerCase();
            const label = (btn.querySelector('[data-tid="Select__label"]') || {}).textContent || '';
            return text.includes('скач') || label.toLowerCase().includes('скач');
        });
        """
    )
    if js_target:
        WebDriverWait(driver, 10).until(lambda _: js_target.is_displayed() and js_target.is_enabled())
        return js_target
    raise TimeoutException("Не удалось найти кнопку 'Скачать'") from last_exc


def _find_original_format_button(driver: webdriver.Chrome) -> WebElement:
    wait = WebDriverWait(driver, 20)
    xpaths = [
        "//div[contains(@class,'react-ui-1mhlayz')]//span[contains(translate(normalize-space(.),'ДОКУМЕНТ В ИСХОДНОМ ФОРМАТЕ','документ в исходном формате'),'документ в исходном формате')]",
        "//button[contains(translate(normalize-space(.),'ДОКУМЕНТ В ИСХОДНОМ ФОРМАТЕ','документ в исходном формате'),'документ в исходном формате')]",
        "//li[contains(translate(normalize-space(.),'ДОКУМЕНТ В ИСХОДНОМ ФОРМАТЕ','документ в исходном формате'),'документ в исходном формате')]",
        "//*[contains(translate(normalize-space(.),'ДОКУМЕНТ В ИСХОДНОМ ФОРМАТЕ','документ в исходном формате'),'документ в исходном формате')]",
    ]
    last_exc: Optional[Exception] = None
    for xp in xpaths:
        try:
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            return element
        except Exception as exc:
            last_exc = exc
    raise TimeoutException("Не найден пункт 'Документ в исходном формате'") from last_exc


def highlight_element(driver: webdriver.Chrome, element: WebElement, color: str) -> None:
    try:
        driver.execute_script(
            "arguments[0].style.transition='background-color 0.3s ease';"
            "arguments[0].style.backgroundColor=arguments[1];"
            "arguments[0].style.outline='2px solid #008000';",
            element,
            color,
        )
    except Exception:
        pass


def wait_for_space_to_close() -> None:
    try:
        log_info("Ожидание закрытия окна Chrome: нажмите пробел + Enter.")
        while True:
            response = input("Нажмите пробел и Enter, чтобы закрыть окно Chrome: ")
            if response == " ":
                break
    except EOFError:
        pass
    except KeyboardInterrupt:
        raise


def main(argv: Optional[List[str]] = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    date_flag_from_kv = False
    lookin_flag_from_kv = False
    remaining_args: List[str] = []
    for arg in argv:
        normalized = arg.strip()
        lowered = normalized.lower()
        if lowered == "date":
            date_flag_from_kv = True
            continue
        if lowered.startswith("date="):
            value = lowered.split("=", 1)[1]
            if value in {"1", "true", "yes"}:
                date_flag_from_kv = True
            continue
        if lowered == "lookin":
            lookin_flag_from_kv = True
            continue
        if lowered.startswith("lookin="):
            value = lowered.split("=", 1)[1]
            if value in {"1", "true", "yes"}:
                lookin_flag_from_kv = True
            continue
        remaining_args.append(arg)

    parser = argparse.ArgumentParser(description="Контур automation")
    parser.add_argument("--date", action="store_true", help="Обновить только поля даты")
    parser.add_argument(
        "--lookin",
        action="store_true",
        help="Отправлять документы в LOOKIN API вместо Supabase",
    )
    args = parser.parse_args(remaining_args)

    date_mode = args.date or date_flag_from_kv
    use_lookin = args.lookin or lookin_flag_from_kv
    try:
        target_inns = fetch_target_inns()
    except Exception as exc:
        log_error("Не удалось получить список ИНН из Supabase: %s", exc)
        return 1

    driver: webdriver.Chrome
    try:
        driver = build_driver()
    except Exception as exc:  # pragma: no cover
        log_error("Не удалось запустить Chrome: %s", exc)
        return 1

    try:
        open_auth_page(driver)
        wait_for_selection_page(driver)
        process_companies(driver, target_inns, date_mode=date_mode, use_lookin=use_lookin)
        wait_for_space_to_close()
    except KeyboardInterrupt:
        log_info("Остановка по запросу пользователя.")
    except Exception as exc:
        log_error("Ошибка при работе с auth.kontur.ru: %s", exc)
        return 1
    finally:
        try:
            driver.quit()
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
