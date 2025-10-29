import os
import sys
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

try:
    from webdriver_manager.chrome import ChromeDriverManager  # type: ignore
except Exception:  # pragma: no cover
    ChromeDriverManager = None  # type: ignore


SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"

STEALTH = os.environ.get("STEALTH", "1")
ROLE_ID = os.environ.get("ROLE_ID", "")
REPO_ROOT = Path(__file__).resolve().parent.parent
LOCAL_CHROMEDRIVER = REPO_ROOT / "bin" / "chromedriver"


def _chromedriver_binary_ok(path: Path) -> bool:
    if not path.exists() or not os.access(path, os.X_OK):
        return False
    try:
        signature = path.open("rb").read(4)
    except OSError:
        return False
    return signature.startswith(b"\x7fELF") or signature.startswith(b"MZ")


def _make_service() -> Service:
    env_path = (os.environ.get("CHROMEDRIVER") or "").strip()
    candidates = [env_path, str(LOCAL_CHROMEDRIVER), "/usr/bin/chromedriver"]
    for raw in candidates:
        if not raw:
            continue
        path = Path(raw)
        if _chromedriver_binary_ok(path):
            print(f"[run] Использую chromedriver: {path}")
            return Service(str(path))
        if path.exists():
            print(f"[run] {path} найден, но не похож на исполняемый chromedriver — пропускаю.")
    if env_path:
        print(f"[run] CHROMEDRIVER={env_path}, но рабочий драйвер не найден. Скачаю новый…")
    if ChromeDriverManager is None:
        raise RuntimeError("Chromedriver not found и webdriver-manager недоступен.")
    return Service(ChromeDriverManager().install())


def build_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    user_data_dir = os.environ.get("USER_DATA_DIR") or str(Path.cwd() / "profile")
    Path(user_data_dir).mkdir(parents=True, exist_ok=True)
    options.add_argument(f"--user-data-dir={user_data_dir}")

    # Headless OFF by default for visible role selection
    headless = os.environ.get("HEADLESS", "0") == "1"
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

    if STEALTH == "1":
        options.add_experimental_option("excludeSwitches", ["enable-automation"])  # hide automation banner
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--lang=ru-RU")

    if os.environ.get("CHROME_BIN"):
        options.binary_location = os.environ["CHROME_BIN"]

    driver = webdriver.Chrome(service=_make_service(), options=options)
    if STEALTH == "1":
        _apply_stealth(driver)
    return driver


def _apply_stealth(driver: webdriver.Chrome) -> None:
    try:
        try:
            driver.execute_cdp_cmd("Network.enable", {})
        except Exception:
            pass
        try:
            ua = driver.execute_script("return navigator.userAgent") or ""
        except Exception:
            ua = ""
        new_ua = ua.replace("HeadlessChrome", "Chrome") if ua else None
        try:
            driver.execute_cdp_cmd(
                "Network.setUserAgentOverride",
                {
                    "userAgent": new_ua
                    or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
                    "platform": "Win32",
                },
            )
        except Exception:
            pass
        try:
            driver.execute_cdp_cmd(
                "Network.setExtraHTTPHeaders",
                {"headers": {"Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"}},
            )
        except Exception:
            pass
        try:
            driver.execute_cdp_cmd("Emulation.setTimezoneOverride", {"timezoneId": "Europe/Moscow"})
        except Exception:
            pass
        try:
            driver.execute_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
                "window.chrome = window.chrome || { runtime: {} };"
                "Object.defineProperty(navigator, 'languages', {get: () => ['ru-RU','ru','en-US','en']});"
                "Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});"
                "Object.defineProperty(navigator, 'platform', {get: () => 'Win32'});"
            )
        except Exception:
            pass
    except Exception:
        pass


def print_roles(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    def highlight_role(role_value: str) -> None:
        if not role_value:
            return
        try:
            driver.execute_script(
                """
                const targetVal = arguments[0];
                if (!targetVal) return;
                const nodes = Array.from(document.querySelectorAll('[name="roleId"]'));
                for (const node of nodes) {
                    const value = node.getAttribute('value') || '';
                    if (value === targetVal) {
                        node.style.outline = '3px solid #6d28d9';
                        node.style.backgroundColor = '#ede9fe';
                        try { node.scrollIntoView({ block: 'center', inline: 'center' }); } catch (e) {}
                        break;
                    }
                }
                """,
                role_value,
            )
        except Exception:
            pass

    try:
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[name="roleId"]')))
    except Exception:
        pass
    try:
        roles = driver.execute_script(
            """
            return Array.from(document.querySelectorAll('[name="roleId"]'))
              .map(el => ({ value: el.getAttribute('value') || '', text: (el.textContent||el.value||'').trim() }));
            """
        ) or []
        if roles:
            print("[roles] Доступные роли:")
            for r in roles:
                v = r.get('value')
                t = r.get('text')
                print(f" - value={v} text={t}")
            highlight_role(ROLE_ID)
        else:
            print("[roles] Не удалось считать список ролей (возможно, WAF/Forbidden в headless)")
    except Exception:
        print("[roles] Не удалось считать список ролей")


def main() -> int:
    print("[run] Открываю страницу выбора роли…")
    try:
        driver = build_driver()
    except Exception as e:
        print(f"[run] Не удалось запустить Chrome: {e}", file=sys.stderr)
        return 2
    wait = WebDriverWait(driver, 20)

    driver.get(BACK_TO_SELECT_ROLE_URL)
    if "/SelectRole" not in driver.current_url:
        try:
            driver.get(SELECT_DEPARTMENT_URL)
        except Exception:
            pass
    time.sleep(0.3)
    print(f"[run] current_url: {driver.current_url}")
    if "/SelectRole" in driver.current_url:
        print_roles(driver, wait)
    else:
        print("[hint] Похоже, не на странице SelectRole (возможно, нужен GUI или активная сессия).")

    print("[wait] Скрипт не будет закрывать окно. Закройте браузер вручную или нажмите Ctrl+C здесь.")
    try:
        while True:
            try:
                handles = driver.window_handles
            except Exception:
                break
            if not handles:
                break
            time.sleep(0.5)
    except KeyboardInterrupt:
        print("\n[wait] Прервано пользователем. Окно останется открытым.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
