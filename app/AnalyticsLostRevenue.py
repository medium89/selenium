from __future__ import annotations

import csv
import datetime as dt
import os
import time
from typing import List, Optional, Tuple

from selenium import webdriver
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait

try:
    from webdriver_manager.chrome import ChromeDriverManager  # type: ignore
except Exception:  # pragma: no cover
    ChromeDriverManager = None  # type: ignore


# ---- URLs / constants ----
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ANALYTICS_URL = (
    "https://officemanager.dodopizza.ru/OfficeManager/Analytics/1098?"
    "native_filters_key=jOLx2lHq962ZqghJfvg7aIgHLScTqAXtQ9ywqU2RLLOdhW3f8GtLVbWtiVinrwMv"
)
ROLE_ID = os.environ.get("ROLE_ID", "7")  # Менеджер проектов

CSV_FILE = os.environ.get("LOST_REVENUE_CSV", "reports/lost_revenue.csv")
STEALTH = os.environ.get("STEALTH", "1")
PORT = int(os.environ.get("PORT", "9225"))


def highlight(driver: webdriver.Chrome, el) -> None:
    try:
        driver.execute_script(
            "arguments[0].style.outline='2px solid #00aa00';"
            "arguments[0].style.backgroundColor='#c6f7c6';"
            "arguments[0].scrollIntoView({block:'center'});",
            el,
        )
    except Exception:
        pass


def _make_service() -> Service:
    path = os.environ.get("CHROMEDRIVER", "/usr/bin/chromedriver")
    if path and os.path.exists(path):
        return Service(path)
    if ChromeDriverManager is None:
        raise RuntimeError("Chromedriver not found and webdriver-manager is unavailable.")
    return Service(ChromeDriverManager().install())


def build_driver() -> webdriver.Chrome:
    def make_options(profile_dir: str) -> webdriver.ChromeOptions:
        opts = webdriver.ChromeOptions()
        os.makedirs(profile_dir, exist_ok=True)
        opts.add_argument(f"--user-data-dir={profile_dir}")
        if os.environ.get("HEADLESS", "1") == "1":
            opts.add_argument("--headless=new")
            opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--window-size=1920,1080")
        if STEALTH == "1":
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])  # hide automation banner
            opts.add_experimental_option("useAutomationExtension", False)
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument("--lang=ru-RU")
        if os.environ.get("CHROME_BIN"):
            opts.binary_location = os.environ["CHROME_BIN"]
        if _wait_port(PORT, 1):
            opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{PORT}")
        return opts

    base_dir = os.environ.get("USER_DATA_DIR") or os.path.join(os.getcwd(), "profile")
    options = make_options(base_dir)
    try:
        driver = webdriver.Chrome(service=_make_service(), options=options)
    except SessionNotCreatedException as e:
        msg = str(e)
        if "user data directory is already in use" in msg or "already in use" in msg:
            # Fallback: use unique temp user-data-dir to avoid profile lock
            uniq_dir = os.path.join(base_dir, f"_run_{int(time.time())}")
            options = make_options(uniq_dir)
            driver = webdriver.Chrome(service=_make_service(), options=options)
        else:
            raise

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


def ensure_role_selected(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    if "/SelectRole" in driver.current_url:
        # dump roles for convenience
        try:
            roles = driver.execute_script(
                "return Array.from(document.querySelectorAll('[name=\"roleId\"]')).map(r=>({v:r.value,t:(r.textContent||r.value||'').trim()}));"
            ) or []
            if roles:
                print("[roles] Доступные роли:")
                for r in roles:
                    print(f"  value={r.get('v')} text={r.get('t')}")
        except Exception:
            pass
        try:
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="roleId"][value="{ROLE_ID}"]'))).click()
        except Exception:
            try:
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'[name="roleId"][value="{ROLE_ID}"]'))).click()
            except Exception:
                pass
        try:
            WebDriverWait(driver, 10).until(lambda d: "/SelectRole" not in d.current_url)
        except Exception:
            pass


def reset_to_select_department(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    # Go back to role selection to drop current department binding
    try:
        driver.get(BACK_TO_SELECT_ROLE_URL)
    except Exception:
        pass
    # If role selection is shown, pick ROLE_ID
    ensure_role_selected(driver, wait)
    # Open department selection explicitly
    try:
        driver.get(SELECT_DEPARTMENT_URL)
    except Exception:
        pass
    # Wait for the list of departments to be present
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'button[name="uuid"], a[name="uuid"]'))
        )
    except Exception:
        pass


def open_select_department(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    # Always ensure we reset to the department selection for each iteration
    reset_to_select_department(driver, wait)


def get_cities(driver: webdriver.Chrome, wait: WebDriverWait) -> List[Tuple[str, str]]:
    open_select_department(driver, wait)
    try:
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'button[name="uuid"], a[name="uuid"]')))
    except Exception:
        raise RuntimeError("Не удалось получить список городов на SelectDepartment")
    try:
        items = driver.execute_script(
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
    seen = set()
    cities: List[Tuple[str, str]] = []
    for it in items:
        uuid = it.get("uuid")
        name = it.get("name")
        if uuid and name and uuid not in seen:
            seen.add(uuid)
            cities.append((name, uuid))
    cities.sort(key=lambda x: x[0].lower())
    return cities


def select_city(driver: webdriver.Chrome, wait: WebDriverWait, city_uuid: str) -> None:
    open_select_department(driver, wait)
    ensure_role_selected(driver, wait)
    js_find = """
        var id = arguments[0];
        var nodes = Array.from(document.querySelectorAll("button[name='uuid'], a[name='uuid']"));
        for (var n of nodes) {
            var v = n.getAttribute('value') || n.getAttribute('data-value') ||
                    n.getAttribute('data-uuid') || n.getAttribute('uuid') ||
                    n.getAttribute('data-id');
            if (v === id) {
                n.scrollIntoView({block:'center'});
                return n;
            }
        }
        return null;
    """
    el = None
    # Try a couple of times in case the list is still rendering
    for _ in range(3):
        try:
            el = driver.execute_script(js_find, city_uuid)
        except Exception:
            el = None
        if el:
            break
    if not el:
        # Fallback: try multiple CSS attribute combinations
        candidates = [
            f"button[name='uuid'][value='{city_uuid}']",
            f"button[name='uuid'][data-value='{city_uuid}']",
            f"button[name='uuid'][data-uuid='{city_uuid}']",
            f"button[name='uuid'][uuid='{city_uuid}']",
            f"button[name='uuid'][data-id='{city_uuid}']",
            f"a[name='uuid'][value='{city_uuid}']",
            f"a[name='uuid'][data-value='{city_uuid}']",
            f"a[name='uuid'][data-uuid='{city_uuid}']",
            f"a[name='uuid'][uuid='{city_uuid}']",
            f"a[name='uuid'][data-id='{city_uuid}']",
        ]
        for sel in candidates:
            try:
                el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                if el:
                    break
            except Exception:
                el = None
    if not el:
        raise RuntimeError(f"Не найден элемент города с uuid={city_uuid}")
    try:
        el.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click()", el)
        except Exception:
            raise
    # Wait until we leave the SelectDepartment page to ensure the switch is applied
    try:
        WebDriverWait(driver, 10).until(lambda d: "/SelectDepartment" not in d.current_url)
    except Exception:
        pass
    


def read_lost_revenue(driver: webdriver.Chrome) -> str:
    import re
    try:
        elems = driver.find_elements(By.CSS_SELECTOR, "*")
    except Exception:
        elems = []
    pattern = re.compile(r"(\d+[\.,]?\d*)\s*%")
    for el in elems:
        try:
            txt = (el.text or "").strip()
        except Exception:
            continue
    
        if not txt:
            continue
        if ("Доля упущ" in txt or "упущенной выруч" in txt) and "%" in txt:
            m = pattern.search(txt.replace("\xa0", " "))
            if m:
                return m.group(1).replace(",", ".") + "%"
            return txt
    return ""


def wait_lost_revenue_update(
    driver: webdriver.Chrome,
    previous: Optional[str] = None,
    timeout: float = 35.0,
    poll: float = 0.6,
) -> str:
    min_wait = time.time() + 2.0
    deadline = time.time() + max(timeout, 1.0)
    last_val: Optional[str] = None
    stable_reads = 0
    while time.time() < deadline:
        current = (read_lost_revenue(driver) or "").strip()
        if not current:
            last_val = None
            stable_reads = 0
            time.sleep(poll)
            continue
        if current == last_val:
            stable_reads += 1
        else:
            last_val = current
            stable_reads = 1
        if stable_reads >= 2:
            if current != (previous or "").strip():
                return current
            if time.time() >= min_wait:
                return current
        time.sleep(poll)
    return last_val or previous or ""


def write_csv_row(path: str, row: List[str]) -> None:
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    exists = os.path.exists(path)
    with open(path, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        if not exists:
            w.writerow(["Город", "Дата", "Доля упущенной выручки"])
        w.writerow(row)




def _wait_port(port: int, timeout: int = 10) -> bool:
    import socket
    for _ in range(timeout * 10):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", port)) == 0:
                return True
        
    return False


def main() -> int:
    driver = build_driver()
    wait = WebDriverWait(driver, 25)
    try:
        # Get all cities, iterate
        cities = get_cities(driver, wait)
        today = dt.date.today().strftime("%d.%m.%Y")
        print(f"[CITIES] Найдено: {len(cities)} — {', '.join([c[0] for c in cities])}")
        for idx, (city_name, city_uuid) in enumerate(cities, start=1):
            print("\n" + "#" * 80)
            print(f"[CITY] ({idx}/{len(cities)}) {city_name}")
            try:
                select_city(driver, wait, city_uuid)
                goto_analytics(driver, wait)
                # Wait base readiness
                try:
                    WebDriverWait(driver, 30).until(
                        lambda d: d.execute_script("return document.readyState") == "complete"
                    )
                except Exception:
                    pass
                # Ensure we're still on analytics before opening filters
                try:
                    ensure_on_analytics_or_retry(driver, wait, retries=2)
                except Exception:
                    # One more forced attempt
                    goto_analytics(driver, wait)
                    ensure_on_analytics_or_retry(driver, wait, retries=1)

                prev_metric_city: Optional[str] = None
                # Попытка применить фильтры (Страна=Россия; Пиццерии=по городу)
                try:
                    open_filter_panel(driver)
                    # Ждём полной готовности панели: селекты отрендерены, спиннеры/скелетоны скрыты
                    wait_filters_loaded(driver, wait, timeout=30)
                    set_country_russia_only(driver)
                    city_key = (city_name or "").strip().lower()
                    if city_key == "ставрополь":
                        pizzerias = get_pizzerias_for_city(driver, city_name) or []
                        if not pizzerias:
                            pizzerias = [city_name]
                        for branch_idx, branch_name in enumerate(pizzerias, start=1):
                            prev_metric_branch: Optional[str] = None
                            if branch_idx > 1:
                                open_filter_panel(driver)
                                wait_filters_loaded(driver, wait, timeout=30)
                            set_country_russia_only(driver)
                            try:
                                selected = set_pizzerias_for_city(
                                    driver, city_name, exact_pizzerias=[branch_name]
                                )
                                if not selected:
                                    raise RuntimeError("пиццерия не была выбрана в фильтре")
                                prev_metric_branch = read_lost_revenue(driver)
                                apply_filters(driver)
                                print(f"[FILTER] Применены фильтры для {branch_name}")
                                branch_val = wait_lost_revenue_update(driver, previous=prev_metric_branch)
                                write_csv_row(CSV_FILE, [branch_name, today, branch_val or "(нет данных)"])
                                print(f"[CSV] {branch_name}: {branch_val}")
                            except Exception as branch_err:
                                print(f"[WARN] Ошибка в пиццерии {branch_name}: {branch_err}")
                                write_csv_row(CSV_FILE, [branch_name, today, f"ОШИБКА: {branch_err}"])
                        continue

                    set_pizzerias_for_city(driver, city_name)
                    prev_metric_city = read_lost_revenue(driver)
                    apply_filters(driver)
                    print(f"[FILTER] Применены фильтры для {city_name}")
                except Exception as e:
                    print(f"[filters] Не удалось применить фильтры: {e}")
                val = wait_lost_revenue_update(driver, previous=prev_metric_city)
                write_csv_row(CSV_FILE, [city_name, today, val or "(нет данных)"])
                print(f"[CSV] {city_name}: {val}")
            except Exception as e:
                print(f"[WARN] Ошибка в городе {city_name}: {e}")
                write_csv_row(CSV_FILE, [city_name, today, f"ОШИБКА: {e}"])
    finally:
        try:
            driver.quit()
        except Exception:
            pass
    print(f"[DONE] Готово! Файл {CSV_FILE} сохранён.")
    return 0


# ---------------- Helpers to operate native filter panel (best-effort) ----------------

def open_filter_panel(driver: webdriver.Chrome) -> None:
    selectors = [
        "[data-test='native-filters-trigger']",
        "button[aria-label='Filters']",
        "button:has(svg[data-icon='filter'])",
        "button[title='Фильтры']",
    ]
    for sel in selectors:
        try:
            btns = driver.find_elements(By.CSS_SELECTOR, sel)
            if btns:
                highlight(driver, btns[0])
                btns[0].click()
                return
        except Exception:
            continue
    try:
        btn = driver.find_element(By.CSS_SELECTOR, "#single-spa-application\\:supersetDashboardPlugin button")
        highlight(driver, btn)
        btn.click()
        
    except Exception:
        pass


def wait_filters_loaded(driver: webdriver.Chrome, wait: WebDriverWait, timeout: int = 30) -> None:
    # Ждём, пока панель фильтров появится, селекты будут отрендерены и пропадут спиннеры/скелетоны (Ant Design)
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
        WebDriverWait(driver, timeout).until(_ready)
    except Exception:
        # Best-effort: proceed even if we timed out; subsequent steps may still work
        pass


def _open_filter_by_label(driver: webdriver.Chrome, label_text: str | List[str]):
    js = """
        var labels = Array.isArray(arguments[0]) ? arguments[0] : [arguments[0]];
        var nodes = Array.from(document.querySelectorAll('*')).filter(n=>{
          var t=(n.textContent||'').trim();
          if(!t) return false;
          var tl = t.toLowerCase();
          return labels.some(lbl => tl.includes(String(lbl||'').toLowerCase()));
        });
        var panel = document.querySelector('[class*="filter"], [data-test="native-filters"]') || document.body;
        for (var i=0;i<nodes.length;i++){
          var el=nodes[i]; if(!panel.contains(el)) continue;
          var root=el;
          for(var j=0;j<6 && root && !root.querySelector('div.ant-select'); j++){ root=root.parentElement; }
          if(root){
            var select = root.querySelector('div.ant-select');
            if(select){ select.click(); return root; }
          }
        }
        return null;
    """
    try:
        return driver.execute_script(js, label_text)
    except Exception:
        return None


def _get_visible_dropdown(driver: webdriver.Chrome):
    js = """
        var lists = Array.from(document.querySelectorAll('.ant-select-dropdown'));
        var visible = lists.filter(function(el){
          var st = window.getComputedStyle(el);
          return st && st.display !== 'none' && st.visibility !== 'hidden' && el.offsetParent !== null;
        });
        return visible.length ? visible[visible.length - 1] : null;
    """
    try:
        return driver.execute_script(js)
    except Exception:
        return None


def _dropdown_wait_options(driver: webdriver.Chrome, timeout: float = 5.0) -> None:
    def _has_opts(drv) -> bool:
        try:
            dd = _get_visible_dropdown(drv)
            if not dd:
                return False
            opts = dd.find_elements(By.CSS_SELECTOR, ".ant-select-item-option")
            return bool(opts)
        except Exception:
            return False
    try:
        WebDriverWait(driver, timeout).until(lambda d: _has_opts(d))
    except Exception:
        pass


def _get_filters_container(driver: webdriver.Chrome):
    # Prefer native filters container, fallback to ant-drawer body
    js = """
        return document.querySelector('[data-test="native-filters"]') ||
               document.querySelector('.ant-drawer .ant-drawer-body') ||
               document.querySelector('.ant-drawer');
    """
    try:
        return driver.execute_script(js)
    except Exception:
        return None


def _clear_tokens_within_select(driver: webdriver.Chrome, select_el) -> None:
    # Aggressively clear all selected tokens inside the given select
    for _ in range(4):  # a few passes in case DOM updates between clicks
        removed_any = False
        # Click individual token close icons
        try:
            close_icons = select_el.find_elements(By.CSS_SELECTOR, ".ant-tag-close-icon, .ant-select-selection-item-remove")
        except Exception:
            close_icons = []
        for btn in close_icons:
            try:
                btn.click(); removed_any = True
            except Exception:
                try:
                    driver.execute_script("arguments[0].click()", btn); removed_any = True
                except Exception:
                    pass
        # Click top-level clear button if present
        try:
            clears = select_el.find_elements(By.CSS_SELECTOR, ".ant-select-clear")
        except Exception:
            clears = []
        for btn in clears:
            try:
                btn.click(); removed_any = True
            except Exception:
                try:
                    driver.execute_script("arguments[0].click()", btn); removed_any = True
                except Exception:
                    pass
        # If nothing was removed this pass, stop
        if not removed_any:
            break


def _open_select_dropdown(driver: webdriver.Chrome, select_el) -> None:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'})", select_el)
    except Exception:
        pass
    try:
        select_el.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click()", select_el)
        except Exception:
            pass


def _close_dropdown(driver: webdriver.Chrome) -> None:
    # Попробовать закрыть выпадающий список: ESC, клик по заголовку панели, blur активного элемента
    try:
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception:
        pass
    try:
        header = driver.find_elements(By.CSS_SELECTOR, ".ant-drawer-header, [data-test='native-filters'] h4")
        if header:
            try:
                header[0].click()
            except Exception:
                driver.execute_script("arguments[0].click()", header[0])
    except Exception:
        pass
    try:
        driver.execute_script("if(document.activeElement) document.activeElement.blur();")
    except Exception:
        pass


def _wait_tokens_containing(select_el, needle: str, timeout: float = 6.0) -> bool:
    import time as _t
    end = _t.time() + timeout
    needle_l = (needle or "").lower()
    while _t.time() < end:
        try:
            tokens = select_el.find_elements(By.CSS_SELECTOR, ".ant-select-selection-item, .ant-tag")
        except Exception:
            tokens = []
        for tk in tokens:
            try:
                txt = (tk.text or "").strip().lower()
            except Exception:
                txt = ""
            if txt and needle_l in txt:
                return True
        try:
            time.sleep(0.2)
        except Exception:
            pass
    return False


def _collect_tokens_text(select_el) -> List[str]:
    try:
        tokens = select_el.find_elements(By.CSS_SELECTOR, ".ant-select-selection-item, .ant-tag")
    except Exception:
        tokens = []
    values: List[str] = []
    for tk in tokens:
        try:
            txt = (tk.text or "").strip()
        except Exception:
            txt = ""
        if txt:
            values.append(txt)
    return values


def _dropdown_type_query(driver: webdriver.Chrome, query: str) -> None:
    dd = _get_visible_dropdown(driver)
    inputs = []
    if dd:
        try:
            inputs = dd.find_elements(By.CSS_SELECTOR, ".ant-select-selection-search-input, input")
        except Exception:
            inputs = []
    # Fallback: искать поле поиска в триггере открытого селекта
    if not inputs:
        try:
            open_sels = driver.find_elements(By.CSS_SELECTOR, ".ant-select.ant-select-open, .ant-select-focused")
        except Exception:
            open_sels = []
        for root in open_sels:
            try:
                cand = root.find_elements(By.CSS_SELECTOR, ".ant-select-selection-search-input, input")
            except Exception:
                cand = []
            if cand:
                inputs = cand
                break
    if not inputs:
        return
    field = inputs[0]
    try:
        field.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click()", field)
        except Exception:
            pass
    # Надёжная очистка и ввод
    try:
        field.clear()
    except Exception:
        pass
    try:
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(Keys.BACK_SPACE).perform()
    except Exception:
        pass
    try:
        field.send_keys(query)
    except Exception:
        try:
            ActionChains(driver).send_keys(query).perform()
        except Exception:
            pass
    _dropdown_wait_options(driver, timeout=6.0)


def _select_options_containing(driver: webdriver.Chrome, needle: str, max_count: int = 25) -> int:
    dd = _get_visible_dropdown(driver)
    if not dd:
        return 0
    try:
        options = dd.find_elements(By.CSS_SELECTOR, ".ant-select-item-option")
    except Exception:
        options = []
    picked = 0
    for opt in options:
        if picked >= max_count:
            break
        try:
            txt = (opt.text or "").strip()
        except Exception:
            continue
        if txt and needle.lower() in txt.lower():
            try:
                opt.click(); picked += 1
            except Exception:
                try:
                    driver.execute_script("arguments[0].click()", opt); picked += 1
                except Exception:
                    pass
    return picked


def _select_first_option_containing(driver: webdriver.Chrome, needle: str) -> bool:
    return _select_options_containing(driver, needle, max_count=1) > 0


def _select_option_exact(driver: webdriver.Chrome, value: str) -> bool:
    dd = _get_visible_dropdown(driver)
    if not dd:
        return False
    try:
        options = dd.find_elements(By.CSS_SELECTOR, ".ant-select-item-option")
    except Exception:
        options = []
    target = (value or "").strip().lower()
    for opt in options:
        try:
            txt = (opt.text or "").strip()
        except Exception:
            txt = ""
        if txt and txt.strip().lower() == target:
            try:
                opt.click()
            except Exception:
                try:
                    driver.execute_script("arguments[0].click()", opt)
                except Exception:
                    return False
            return True
    return False


def _select_options_starting_with(driver: webdriver.Chrome, prefix: str, max_count: int = 25) -> int:
    dd = _get_visible_dropdown(driver)
    if not dd:
        return 0
    try:
        options = dd.find_elements(By.CSS_SELECTOR, ".ant-select-item-option")
    except Exception:
        options = []
    picked = 0
    p = (prefix or "").strip().lower()
    for opt in options:
        if picked >= max_count:
            break
        try:
            txt = (opt.text or "").strip().lower()
        except Exception:
            continue
        if not txt:
            continue
        # Подбираем только начинающиеся с префикса (например, "абинск-1" при запросе "абинск")
        if txt.startswith(p):
            try:
                opt.click(); picked += 1
            except Exception:
                try:
                    driver.execute_script("arguments[0].click()", opt); picked += 1
                except Exception:
                    pass
    return picked


def _find_select_by_label(driver: webdriver.Chrome, labels: List[str]):
    js = """
        var labels = (arguments[0]||[]).map(x=>String(x||'').toLowerCase());
        var items = Array.from(document.querySelectorAll('form.ant-form .ant-row.ant-form-item'));
        for (var it of items){
            var h = it.querySelector('h4');
            var t = h && (h.textContent||'').trim().toLowerCase();
            if (!t) continue;
            if (labels.some(lb => t.includes(lb))) {
                var sel = it.querySelector('div.ant-select');
                if (sel) return sel;
            }
        }
        return null;
    """
    try:
        return driver.execute_script(js, labels)
    except Exception:
        return None


def _clear_all_but_keep(driver: webdriver.Chrome, select_el, keep_texts: List[str]) -> None:
    want = {t.lower(): True for t in keep_texts}
    try:
        tokens = select_el.find_elements(By.CSS_SELECTOR, ".ant-select-selection-item")
    except Exception:
        tokens = []
    for tk in tokens:
        try:
            txt_el = tk.find_element(By.CSS_SELECTOR, ".tag-content")
            txt = (txt_el.text or "").strip()
        except Exception:
            txt = ""
        if not txt or txt.lower() not in want:
            try:
                btn = tk.find_element(By.CSS_SELECTOR, ".ant-tag-close-icon, .ant-select-selection-item-remove")
                btn.click()
            except Exception:
                try:
                    driver.execute_script("arguments[0].click()", btn)
                except Exception:
                    pass


def set_country_russia_only(driver: webdriver.Chrome) -> None:
    # Find country select by label and leave only 'Россия'
    sel = _find_select_by_label(driver, ["страна", "country"])
    # Не трогаем другие фильтры (например, "Период"): без строгого совпадения по метке — выходим
    if not sel:
        return
    # remove all except Russia if present
    _clear_all_but_keep(driver, sel, ["Россия"]) 
    # ensure Russia present; if not, search and add
    try:
        # check if token Russia exists
        has_ru = False
        tokens = sel.find_elements(By.CSS_SELECTOR, ".ant-select-selection-item .tag-content")
        for t in tokens:
            try:
                if (t.text or "").strip().lower() == "россия":
                    has_ru = True; break
            except Exception:
                continue
        if not has_ru:
            _open_select_dropdown(driver, sel)
            _dropdown_type_query(driver, "Россия")
            _select_first_option_containing(driver, "Россия")
    except Exception:
        pass
    # Закрыть дропдаун и потерять фокус, чтобы React зафиксировал изменения
    _close_dropdown(driver)


def set_pizzerias_for_city(
    driver: webdriver.Chrome,
    city_name: str,
    exact_pizzerias: Optional[List[str]] = None,
) -> List[str]:
    # Find pizzeria select by label and fill by city pattern
    sel = _find_select_by_label(driver, ["пиццерия", "store", "pizzeria", "пиццер"]) 
    # Не трогаем другие фильтры — без явной метки «Пиццерия» выходим
    if not sel:
        return []
    # clear all current tokens
    _clear_tokens_within_select(driver, sel)

    selected: List[str] = []
    # Если передан точный список пиццерий — выбираем каждую по отдельности
    if exact_pizzerias:
        for name in exact_pizzerias:
            query = name or city_name
            _open_select_dropdown(driver, sel)
            _dropdown_type_query(driver, query)
            found = _select_option_exact(driver, name)
            if not found:
                found = _select_options_starting_with(driver, name, max_count=1) > 0
            if not found:
                found = _select_options_containing(driver, name, max_count=1) > 0
            _close_dropdown(driver)
            if not found:
                continue
            if not _wait_tokens_containing(sel, name, timeout=6.0):
                continue
        selected = _collect_tokens_text(sel)
        return selected

    # open dropdown and type city для массового выбора по городу
    _open_select_dropdown(driver, sel)
    _dropdown_type_query(driver, city_name)
    # Выбираем только варианты, начинающиеся с названия города (исключаем вложенные вхождения типа «алабинск»)
    picked = _select_options_starting_with(driver, city_name, max_count=25)
    # Закрываем выпадающий список для фиксации изменений
    _close_dropdown(driver)
    # Убедиться, что хотя бы один токен добавился; если нет — пробуем ещё раз
    if picked == 0 or not _wait_tokens_containing(sel, city_name, timeout=6.0):
        _open_select_dropdown(driver, sel)
        _dropdown_type_query(driver, city_name)
        _select_options_containing(driver, city_name, max_count=25)
        _close_dropdown(driver)
    selected = _collect_tokens_text(sel)
    return selected


def get_pizzerias_for_city(driver: webdriver.Chrome, city_name: str) -> List[str]:
    sel = _find_select_by_label(driver, ["пиццерия", "store", "pizzeria", "пиццер"])
    if not sel:
        return []
    _open_select_dropdown(driver, sel)
    _dropdown_type_query(driver, city_name)
    dd = _get_visible_dropdown(driver)
    names: List[str] = []
    if dd:
        try:
            options = dd.find_elements(By.CSS_SELECTOR, ".ant-select-item-option")
        except Exception:
            options = []
        city_lower = (city_name or "").strip().lower()
        for opt in options:
            try:
                txt = (opt.text or "").strip()
            except Exception:
                txt = ""
            if not txt:
                continue
            if city_lower and not txt.lower().startswith(city_lower):
                continue
            names.append(txt)
    _close_dropdown(driver)
    unique: List[str] = []
    seen = set()
    for name in names:
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        unique.append(name)
    return unique


def apply_filters(driver: webdriver.Chrome) -> None:
    # Ищем именно кнопку «Применить» внутри панели фильтров и ждём, пока она станет активной
    # Перед нажатием закрываем возможные открытые дропдауны, чтобы клик не перехватывался
    _close_dropdown(driver)
    def _find_apply():
        # 1) Прямо по классу, где бы ни был
        try:
            btns = driver.find_elements(By.CSS_SELECTOR, ".filter-apply-button")
        except Exception:
            btns = []
        visible = []
        for b in btns:
            try:
                if b.is_displayed():
                    visible.append(b)
            except Exception:
                continue
        # Предпочтём ту, где есть текст «Применить»
        for b in visible:
            try:
                t = (b.text or "").strip()
                if "Применить" in t or "Apply" in t:
                    return b
            except Exception:
                continue
        if visible:
            return visible[0]

        # 2) По тексту среди всех кнопок
        try:
            all_btns = driver.find_elements(By.CSS_SELECTOR, "button, [role='button']")
        except Exception:
            all_btns = []
        candidates = []
        for b in all_btns:
            try:
                if not b.is_displayed():
                    continue
                t = (b.text or "").strip()
                if not t:
                    continue
                if "Применить" in t or "Apply" in t:
                    candidates.append(b)
            except Exception:
                continue
        # Отфильтруем, если есть type=submit
        prio = []
        for b in candidates:
            try:
                ty = (b.get_attribute("type") or "").lower()
                cls = (b.get_attribute("class") or "").lower()
            except Exception:
                ty = ""; cls = ""
            score = 0
            if ty == "submit": score += 2
            if "filter-apply-button" in cls: score += 3
            if "superset-button-primary" in cls: score += 1
            prio.append((score, b))
        if prio:
            prio.sort(key=lambda x: x[0], reverse=True)
            return prio[0][1]

        # 3) Попытка найти через форму панели фильтров
        try:
            form = driver.find_element(By.CSS_SELECTOR, "form.ant-form")
            try:
                b = form.find_element(By.CSS_SELECTOR, "button[type='submit'], .filter-apply-button")
                return b
            except Exception:
                pass
        except Exception:
            pass
        return None

    btn = _find_apply()
    if not btn:
        return
    try:
        highlight(driver, btn)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'})", btn)
        # Явно покрасим кнопку и заменим текст, чтобы убедиться, что нашли её
        driver.execute_script(
            "var b=arguments[0];\n"
            "b.style.backgroundColor='#00c853'; b.style.color='#fff'; b.style.borderColor='#00a844';\n"
            "var s=b.querySelector('span'); if(s){ s.textContent='ага, вот эта кнопка'; } else { b.textContent='ага, вот эта кнопка'; }",
            btn,
        )
    except Exception:
        pass

    # Дождаться, когда кнопка не будет disabled
    def _enabled(el) -> bool:
        try:
            # Сначала честная проверка Selenium
            if hasattr(el, "is_enabled"):
                if not el.is_enabled():
                    return False
            # Затем явные атрибуты/классы AntD
            dis_attr = el.get_attribute("disabled")
            if dis_attr is not None and str(dis_attr).lower() in ("", "true", "disabled"):
                return False
            aria = (el.get_attribute("aria-disabled") or "").lower()
            if aria in ("true", "disabled"):
                return False
            cls = (el.get_attribute("class") or "").lower()
            if "ant-btn-disabled" in cls:
                return False
            return True
        except Exception:
            return True

    t0 = time.time()
    # Ждём активации не более 20 секунд (иногда АнтД дольше синхронизируется после выбора)
    while time.time() - t0 < 20:
        if _enabled(btn):
            break
        time.sleep(0.25)

    # Клик по активной кнопке
    try:
        if _enabled(btn):
            try:
                btn.click()
            except Exception:
                # Попробуем клик «мышкой» через ActionChains по центру элемента
                ActionChains(driver).move_to_element(btn).pause(0.05).click().perform()
        else:
            raise Exception("apply-disabled")
        return
    except Exception:
        # Если всё ещё disabled — пробуем сабмит формы напрямую (обход React-disabled)
        try:
            driver.execute_script(
                "var b=arguments[0];\n"
                "if(b){ b.removeAttribute('disabled'); b.classList.remove('ant-btn-disabled'); }\n"
                "b && b.dispatchEvent(new MouseEvent('click',{bubbles:true,composed:true}));\n"
                "var f=b && b.closest('form');\n"
                "if(f){ f.dispatchEvent(new Event('submit',{bubbles:true,cancelable:true})); }",
                btn,
            )
            return
        except Exception:
            pass


def ensure_on_analytics_or_retry(driver: webdriver.Chrome, wait: WebDriverWait, retries: int = 2) -> None:
    for _ in range(max(1, retries + 1)):
        try:
            cur = driver.current_url
        except Exception:
            cur = ""
        if "/OfficeManager/Analytics" in cur:
            return
        # Attempt to recover
            try:
                goto_analytics(driver, wait)
            except Exception:
                pass
    # Final check
    try:
        cur = driver.current_url
    except Exception:
        cur = ""
    if "/OfficeManager/Analytics" not in cur:
        raise RuntimeError("Отклонение с Analytics: не удалось удержаться на странице аналитики")


def goto_analytics(driver: webdriver.Chrome, wait: WebDriverWait, tries: int = 3) -> None:
    # Ensure we land on /OfficeManager/Analytics even if site redirects to OperationalStatistics
    targets = [ANALYTICS_URL.rstrip('/'), ANALYTICS_URL]
    for _ in range(tries):
        for url in targets:
            try:
                driver.get(url)
            except Exception:
                pass
            try:
                WebDriverWait(driver, 20).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
            except Exception:
                pass
            try:
                cur = driver.current_url
            except Exception:
                cur = ""
            if "/OfficeManager/Analytics" in cur:
                return
            # Hard replace to override SPA redirects
            try:
                driver.execute_script("window.stop && window.stop();")
            except Exception:
                pass
            try:
                driver.execute_script("window.location.replace(arguments[0])", url)
            except Exception:
                pass
            try:
                WebDriverWait(driver, 10).until(lambda d: "/OfficeManager/Analytics" in d.current_url)
                return
            except Exception:
                pass
        # Fallback: try to click a visible navigation item labeled 'Аналитика'
        try:
            candidates = driver.find_elements(By.CSS_SELECTOR, "a,button,[role='link'],[role='menuitem']")
        except Exception:
            candidates = []
        for el in candidates:
            try:
                t = (el.text or "").strip()
            except Exception:
                continue
            if t and "Аналитик" in t:
                try:
                    highlight(driver, el)
                except Exception:
                    pass
                try:
                    el.click()
                    WebDriverWait(driver, 10).until(lambda d: "/OfficeManager/Analytics" in d.current_url)
                    return
                except Exception:
                    continue
    raise RuntimeError("Не удалось перейти на страницу Аналитики (перенаправляет на OperationalStatistics)")


if __name__ == "__main__":
    raise SystemExit(main())
