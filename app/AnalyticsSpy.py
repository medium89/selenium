from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from typing import Optional, List, Tuple

import os
import time
import datetime


# URLs/const
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ANALYTICS_URL = (
    "https://officemanager.dodopizza.ru/OfficeManager/Analytics/1098?"
    "native_filters_key=jOLx2lHq962ZqghJfvg7aIgHLScTqAXtQ9ywqU2RLLOdhW3f8GtLVbWtiVinrwMv"
)

# Defaults
ROLE_ID = "7"  # Менеджер проектов
PORT = 9224
STEALTH = os.environ.get("STEALTH", "1")


class AnalyticsSpy:
    def __init__(self):
        self.driver: Optional[webdriver.Chrome] = None
        self.wait: Optional[WebDriverWait] = None

    # ---------- Driver bootstrap ----------
    def _make_service(self) -> Service:
        path = os.environ.get("CHROMEDRIVER", "/usr/bin/chromedriver")
        if path and os.path.exists(path):
            return Service(path)
        return Service(ChromeDriverManager().install())

    def connect_driver(self) -> None:
        options = webdriver.ChromeOptions()
        # profile defaults to $PWD/profile
        user_dir = os.environ.get("USER_DATA_DIR") or os.path.join(os.getcwd(), "profile")
        try:
            os.makedirs(user_dir, exist_ok=True)
        except Exception:
            pass
        options.add_argument(f"--user-data-dir={user_dir}")

        # headless by default
        if os.environ.get("HEADLESS", "1") == "1":
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

        # Prefer attaching to an already-opened Chrome with remote debugging
        if self._wait_port(PORT, 1):
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{PORT}")
            self.driver = webdriver.Chrome(service=self._make_service(), options=options)
        else:
            self.driver = webdriver.Chrome(service=self._make_service(), options=options)

        self.wait = WebDriverWait(self.driver, 25)
        if STEALTH == "1":
            self._apply_stealth()

    def _apply_stealth(self) -> None:
        try:
            try:
                self.driver.execute_cdp_cmd("Network.enable", {})
            except Exception:
                pass
            try:
                ua = self.driver.execute_script("return navigator.userAgent") or ""
            except Exception:
                ua = ""
            new_ua = ua.replace("HeadlessChrome", "Chrome") if ua else None
            try:
                self.driver.execute_cdp_cmd(
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
                self.driver.execute_cdp_cmd(
                    "Network.setExtraHTTPHeaders",
                    {"headers": {"Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"}},
                )
            except Exception:
                pass
            try:
                self.driver.execute_cdp_cmd("Emulation.setTimezoneOverride", {"timezoneId": "Europe/Moscow"})
            except Exception:
                pass
            try:
                self.driver.execute_script(
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

    # ---------- Auth/navigation ----------
    def choose_role(self) -> None:
        # Print roles to help debug
        try:
            roles = self.driver.execute_script(
                """
                return Array.from(document.querySelectorAll('[name="roleId"]'))
                  .map(el => ({ value: el.getAttribute('value') || '', text: (el.textContent||el.value||'').trim() }));
                """
            ) or []
            if roles:
                print("[roles] Доступные роли:")
                for r in roles:
                    print(f"    value={r.get('value')} text={r.get('text')}")
        except Exception:
            pass
        try:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="roleId"][value="{ROLE_ID}"]'))).click()
        except Exception:
            try:
                self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'[name="roleId"][value="{ROLE_ID}"]'))).click()
            except Exception:
                pass
        try:
            WebDriverWait(self.driver, 10).until(lambda d: "/SelectRole" not in d.current_url)
        except Exception:
            pass

    def open_select_department(self) -> None:
        self.driver.get(SELECT_DEPARTMENT_URL)
        if "/SelectRole" in self.driver.current_url:
            self.choose_role()
            # re-open in case of redirects
            try:
                self.driver.get(SELECT_DEPARTMENT_URL)
            except Exception:
                pass

    def get_cities(self) -> List[Tuple[str, str]]:
        self.open_select_department()
        try:
            self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'button[name="uuid"], a[name="uuid"]')))
        except Exception:
            raise RuntimeError("Не удалось дождаться списка городов — возможно, требуется GUI/авторизация")
        try:
            items = self.driver.execute_script(
                """
                return Array.from(document.querySelectorAll('button[name="uuid"], a[name="uuid"]'))
                  .map(b => ({
                    name: (b.textContent || '').trim(),
                    uuid: b.getAttribute('value') || b.getAttribute('data-value') || b.getAttribute('data-uuid') ||
                          b.getAttribute('uuid') || b.getAttribute('data-id') || '',
                  }))
                  .filter(x => x.name && x.uuid);
                """
            ) or []
        except Exception:
            items = []
        cities: List[Tuple[str, str]] = []
        seen = set()
        for it in items:
            uuid = it.get("uuid")
            name = it.get("name")
            if uuid and name and uuid not in seen:
                seen.add(uuid)
                cities.append((name, uuid))
        cities.sort(key=lambda x: x[0].lower())
        if not cities:
            raise RuntimeError("Список городов пуст")
        return cities

    def select_city(self, city_uuid: str) -> None:
        self.open_select_department()
        try:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="uuid"][value="{city_uuid}"]'))).click()
        except Exception:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'a[name="uuid"][value="{city_uuid}"]'))).click()
        time.sleep(0.2)

    # ---------- Filter panel helpers ----------
    def open_filter_panel(self) -> None:
        selectors = [
            "[data-test='native-filters-trigger']",
            "button[aria-label='Filters']",
            "button[title='Фильтры']",
            "button:has(svg[data-icon='filter'])",
        ]
        for sel in selectors:
            try:
                btns = self.driver.find_elements(By.CSS_SELECTOR, sel)
                if btns:
                    btns[0].click(); time.sleep(0.3)
                    return
            except Exception:
                continue
        # Fallback: click first button inside superset app container
        try:
            cont = self.driver.find_element(By.CSS_SELECTOR, "#single-spa-application\\:supersetDashboardPlugin")
            btn = cont.find_element(By.CSS_SELECTOR, "button")
            btn.click(); time.sleep(0.3)
        except Exception:
            pass

    def wait_filters_loaded(self, timeout: int = 30) -> None:
        def _ready(drv) -> bool:
            js = """
                var panel = document.querySelector('[data-test="native-filters"]') ||
                            document.querySelector('.ant-drawer, .ant-drawer-content, .ant-drawer-body');
                if(!panel) return false;
                var selects = panel.querySelectorAll('div.ant-select');
                var spinners = panel.querySelectorAll('.ant-skeleton, .ant-spin-spinning');
                var anyVisible = Array.from(spinners).some(function(el){
                    var st = window.getComputedStyle(el);
                    return st && st.display !== 'none' && st.visibility !== 'hidden' && el.offsetParent !== null;
                });
                return selects.length > 0 && !anyVisible;
            """
            try:
                return bool(drv.execute_script(js))
            except Exception:
                return False
        try:
            WebDriverWait(self.driver, timeout).until(_ready)
        except Exception:
            pass

    # ---------- Stage 1: fetch HTML ----------
    def fetch_and_dump(self) -> None:
        self.driver.get(ANALYTICS_URL)
        # try to wait the framework a bit
        try:
            WebDriverWait(self.driver, 30).until(lambda d: d.execute_script("return document.readyState") == "complete")
        except Exception:
            pass
        # Open filters and ensure they are rendered for the dump
        self.open_filter_panel()
        self.wait_filters_loaded(timeout=30)
        # Extra pause to ensure filters/widgets fully load before dump
        time.sleep(5)
        html = self.driver.page_source or ""
        out_path = os.path.join(os.getcwd(), "spizdil.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"[dump] Сохранено: {out_path} ({len(html)} символов)")

    # ---------- Main ----------
    def run(self) -> int:
        self.connect_driver()
        # choose city interactively
        cities = self.get_cities()
        print("Доступные города:")
        for i, (name, _) in enumerate(cities, start=1):
            print(f" {i}. {name}")
        choice = 1
        try:
            raw = input("Номер города (по умолчанию 1): ").strip()
            if raw.isdigit():
                v = int(raw)
                if 1 <= v <= len(cities):
                    choice = v
        except EOFError:
            pass
        city_name, city_uuid = cities[choice - 1]
        print(f"[city] Выбран: {city_name}")
        self.select_city(city_uuid)
        self.fetch_and_dump()
        return 0

    def close(self) -> None:
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass

    @staticmethod
    def _wait_port(port: int, timeout: int = 10) -> bool:
        import socket as _s
        for _ in range(timeout * 10):
            with _s.socket() as s:
                if s.connect_ex(("127.0.0.1", port)) == 0:
                    return True
            time.sleep(0.1)
        return False


if __name__ == "__main__":
    app = AnalyticsSpy()
    try:
        raise SystemExit(app.run())
    finally:
        app.close()
