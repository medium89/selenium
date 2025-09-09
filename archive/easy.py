# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

import socket, subprocess, time, os

URL = "https://officemanager.dodopizza.ru/OfficeManager/Debiting/PrepareExcelReport"
HOST = "officemanager.dodopizza.ru"

# --- Твой профиль Chrome + remote debugging ---
CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
USER_DATA_DIR = r"C:\Users\ivano\AppData\Local\Google\Chrome\User Data"
PROFILE_DIR = "Default"
DEBUG_PORT = 9222


def is_port_open(host: str, port: int, timeout: float = 0.5) -> bool:
    s = socket.socket()
    s.settimeout(timeout)
    try:
        s.connect((host, port))
        return True
    except Exception:
        return False
    finally:
        try:
            s.close()
        except Exception:
            pass


def ensure_chrome_debug_session() -> None:
    if is_port_open("127.0.0.1", DEBUG_PORT):
        return
    chrome_exe = next((p for p in CHROME_PATHS if os.path.exists(p)), None)
    if not chrome_exe:
        raise FileNotFoundError("Не найден chrome.exe. Проверь CHROME_PATHS.")
    launch_cmd = [
        chrome_exe,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={USER_DATA_DIR}",
        f"--profile-directory={PROFILE_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
    ]
    subprocess.Popen(launch_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(50):
        if is_port_open("127.0.0.1", DEBUG_PORT):
            return
        time.sleep(0.1)
    raise RuntimeError("Chrome не запустился с remote debugging.")


def robust_open_url(driver: webdriver.Chrome, url: str, wait: WebDriverWait) -> None:
    print("URL:", driver.current_url)
    print("Title:", driver.title)
    print("Handles:", driver.window_handles)
    """Принудительно откроем URL: пробуем ВСЕ вкладки, потом создаём новую через CDP и активируем её."""
    # 1) Пытаемся в каждой существующей вкладке
    for h in list(driver.window_handles):
        driver.switch_to.window(h)
        try:
            driver.get(url)
            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            if HOST in (driver.current_url or "").lower():
                return
        except Exception:
            continue

    # 2) Создаём новый target сразу с URL и активируем его
    try:
        tgt = driver.execute_cdp_cmd("Target.createTarget", {"url": url})
        # Активировать (поднять на передний план)
        driver.execute_cdp_cmd("Target.activateTarget", {"targetId": tgt.get("targetId")})
        # Переключаемся на последний handle
        last = driver.window_handles[-1]
        driver.switch_to.window(last)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        if HOST not in (driver.current_url or "").lower():
            # ещё раз жёстко
            driver.execute_script("window.location.href = arguments[0];", url)
            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    except Exception:
        # Фолбэк: создаём обычную новую вкладку и идём туда
        driver.switch_to.new_window('tab')
        driver.get(url)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    # sanity: JS должен работать, и мы на нужном хосте
    assert driver.execute_script("return 2+2") == 4, "JS не исполняется в текущем табе"
    assert HOST in (driver.current_url or "").lower(), f"Ожидали {HOST}, сейчас: {driver.current_url}"


def switch_to_frame_having(driver, locator, max_depth=4) -> bool:
    driver.switch_to.default_content()
    def dfs(depth=0) -> bool:
        if driver.find_elements(*locator):
            return True
        if depth >= max_depth:
            return False
        for f in driver.find_elements(By.TAG_NAME, "iframe"):
            driver.switch_to.frame(f)
            if dfs(depth + 1):
                return True
            driver.switch_to.parent_frame()
        return False
    return dfs(0)


# ---------- MAIN ----------
ensure_chrome_debug_session()

options = Options()
options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    wait = WebDriverWait(driver, 30)

    # НАДЁЖНО открываем нужный адрес (без «Новой вкладки»)
    robust_open_url(driver, URL, wait)

    # Если страница требует даты — быстро проставим (без паники, в try)
    try:
        driver.switch_to.default_content()
        wait.until(EC.element_to_be_clickable((By.ID, "StartDate"))).click()
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "1"))).click()
        driver.switch_to.default_content()
        wait.until(EC.element_to_be_clickable((By.ID, "EndDate"))).click()
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "5"))).click()
    except Exception:
        driver.switch_to.default_content()

    # Ищем кнопку и кликаем (учитывая iframe)
    btn_locator = (By.NAME, "reportButton")
    if not switch_to_frame_having(driver, btn_locator, max_depth=5):
        driver.switch_to.default_content()

    btn = WebDriverWait(driver, 20).until(EC.element_to_be_clickable(btn_locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    try:
        ActionChains(driver).move_to_element(btn).pause(0.1).click(btn).perform()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", btn)

    # На всякий — прямой вызов обработчика
    try:
        driver.execute_script("if (typeof buildReport==='function'){ buildReport(); }")
    except Exception:
        pass

    print("Клик по 'Построить' выполнен.")
    input("Нажми Enter, чтобы закрыть драйвер (Chrome останется открыт)...")

finally:
    driver.quit()
