import os
import json
import sys
import time
import csv
from pathlib import Path
from glob import glob

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

try:
    # Optional: use webdriver-manager when no system chromedriver is provided
    from webdriver_manager.chrome import ChromeDriverManager  # type: ignore
except Exception:  # pragma: no cover
    ChromeDriverManager = None  # type: ignore


# =========================
# Defaults pinned in script
# =========================

# Страницы Office Manager
LOGIN_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
SELECT_DEPARTMENT_URL = LOGIN_URL
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
REPORT_URL = "https://officemanager.dodopizza.ru/OfficeManager/MaterialConsumption"
ROLE_ID = "7"


def env_bool(name: str, default: bool = False) -> bool:
    val = os.environ.get(name)
    if val is None:
        return default
    return str(val).strip().lower() in {"1", "true", "yes", "y", "on"}


def cleanup_profile_locks(user_data_dir: Path) -> None:
    try:
        targets = []
        for sub in (user_data_dir, user_data_dir / "Default"):
            for pattern in ("Singleton*", "DevToolsActivePort"):
                targets.extend(Path(p) for p in glob(str(sub / pattern)))
        for p in targets:
            try:
                if p.is_file() or p.is_symlink():
                    p.unlink(missing_ok=True)
            except Exception:
                pass
    except Exception:
        pass


def build_chrome(headless: bool, user_data_dir: Path) -> webdriver.Chrome:
    options = Options()

    # Attach existing profile to persist auth between runs
    options.add_argument(f"--user-data-dir={str(user_data_dir)}")
    # Ensure the default profile folder exists to avoid first-run popups
    (user_data_dir / "Default").mkdir(parents=True, exist_ok=True)

    # Best-effort: remove stale lock files from a mounted profile
    cleanup_profile_locks(user_data_dir)

    # Headless toggle
    if headless:
        # Modern headless mode
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")

    # Common sensible defaults (esp. for containers/CI)
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

    # Soften Selenium fingerprint in visible mode
    try:
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument("--disable-blink-features=AutomationControlled")
    except Exception:
        pass

    # Allow overriding Chromium binary location if provided
    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin:
        options.binary_location = chrome_bin

    # Prefer system chromedriver (e.g., from Docker image); fallback to manager
    driver_path = os.environ.get("CHROMEDRIVER")
    service: Service
    if driver_path and Path(driver_path).exists():
        service = Service(executable_path=driver_path)
    else:
        if ChromeDriverManager is None:
            raise RuntimeError(
                "Chromedriver not found and webdriver-manager is unavailable."
            )
        # Note: this may require network access on first run
        service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=options)

    # Hide navigator.webdriver and inject a top banner on every page
    try:
        banner_text = os.environ.get("BANNER_TEXT", "js работает")
        anti_detect_js = (
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
            "window.chrome = window.chrome || {};"
            "window.chrome.app = window.chrome.app || {IsInstalled: false};"
            "Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});"
            "Object.defineProperty(navigator, 'languages', {get: () => ['ru-RU','ru','en-US','en']});"
            "try{const w=window; if(w.navigator && 'userAgentData' in w.navigator){Object.defineProperty(w.navigator,'userAgentData',{get:()=>undefined});}}catch(e){};"
        )
        banner_js = (
            "(function(){try{var id='__om_banner__';var t="
            + json.dumps(banner_text)
            + ";var e=document.getElementById(id);if(!e){e=document.createElement('div');e.id=id;e.textContent=t;e.setAttribute('aria-live','polite');e.style.cssText='position:fixed;top:0;left:0;width:100%;height:auto;z-index:2147483647;background:linear-gradient(90deg,#111,#444);color:#fff;text-align:center;font:600 14px/32px -apple-system,system-ui,Segoe UI,Roboto,Arial,sans-serif;letter-spacing:.3px;box-shadow:0 2px 6px rgba(0,0,0,.25);pointer-events:none;';document.documentElement.appendChild(e);}else{e.textContent=t;}}catch(_e){}})();"
        )
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": anti_detect_js + banner_js},
        )
    except Exception:
        pass

    return driver


def main() -> int:
    CSV_FILE = "reports/office.csv"
    csv_path = Path(CSV_FILE)
    try:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

    user_data_dir = Path(os.environ.get("USER_DATA_DIR", str(Path.cwd() / "profile")))
    user_data_dir.mkdir(parents=True, exist_ok=True)

    first_run_marker = user_data_dir / ".auth_initialized"
    desired_headless = env_bool("HEADLESS", True)

    if not first_run_marker.exists() and not desired_headless:
        print("[auth] First run detected — launching Chrome with GUI for login…", flush=True)
        try:
            driver = build_chrome(headless=False, user_data_dir=user_data_dir)
        except Exception as e:
            print(f"[auth] Failed to launch GUI Chrome: {e}", file=sys.stderr)
            return 1
        try:
            driver.get(SELECT_DEPARTMENT_URL)
            print("[auth] Please complete login in the opened browser window.")
            print("[auth] When finished, close the browser window or press ENTER here.")
            start_time = time.time()
            while True:
                if not driver.window_handles:
                    print("[auth] Browser window closed. Proceeding…")
                    break
                if sys.stdin in select_readable():
                    _ = sys.stdin.readline()
                    print("[auth] ENTER received. Proceeding…")
                    break
                if time.time() - start_time > 1800:
                    print("[auth] Timeout reached (30m). Proceeding…")
                    break
                time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except Exception:
                pass
        first_run_marker.write_text("ok", encoding="utf-8")
        print("[auth] First-run auth completed. Marker written.")

    headless_mode = desired_headless
    print(f"[run] Launching Chrome (headless={headless_mode}) with profile at {user_data_dir}")

    try:
        driver = build_chrome(headless=headless_mode, user_data_dir=user_data_dir)
    except Exception as e:
        print(f"[run] Failed to launch Chrome: {e}", file=sys.stderr)
        return 2

    try:
        runner = OfficeManagerRunner(driver=driver, csv_path=csv_path)
        return runner.run()
    finally:
        try:
            driver.quit()
        except Exception:
            pass


def select_readable():
    """Return a list of file-like objects ready for reading (stdin only)."""
    try:
        import select

        rlist, _, _ = select.select([sys.stdin], [], [], 0)
        return rlist
    except Exception:
        # On platforms where select on stdin is unsupported, just return empty
        return []


 


def dump_table_to_csv(driver: webdriver.Chrome, csv_path: Path) -> int:
    """Достаёт таблицу через JS одним снимком (устойчиво к stale-элементам).

    Алгоритм:
      - Берём #report table (или первый table) и собираем матрицу строк через JS
      - Определяем строку заголовка с датами (где есть цифры)
      - Числовые колонки = те, у кого в заголовке есть цифры (иначе последние 5)
      - Пропускаем строку 'Ингредиент' и пустые
      - Пишем: первая текстовая колонка (наименование) + числовые колонки
    """
    matrix = []
    # Несколько попыток на случай перерисовки отчёта
    for _ in range(5):
        try:
            res = driver.execute_script(
                """
                var t = document.querySelector('#report table') || document.querySelector('table');
                if(!t) return [];
                return Array.from(t.querySelectorAll('tr')).map(tr =>
                  Array.from(tr.querySelectorAll('th,td')).map(c => (c.textContent||'').trim())
                );
                """
            )
            matrix = res or []
            if matrix and len(matrix) > 1:
                break
        except Exception:
            pass
        time.sleep(0.1)
    if not matrix:
        return 0

    header = matrix[0]
    num_cols = [i for i, h in enumerate(header) if any(ch.isdigit() for ch in (h or ""))]
    if not num_cols and header:
        num_cols = list(range(max(0, len(header) - 5), len(header)))

    # Начало данных (пропустим подзаголовок 'Ингредиент')
    start_data = 1
    if len(matrix) > 1 and any("ингредиент" in (x or "").lower() for x in matrix[1]):
        start_data = 2

    written = 0
    with csv_path.open("a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        # Заголовок отделения (даты)
        if num_cols:
            hdr = ["НАИМЕНОВАНИЕ"] + [header[i] if i < len(header) else f"COL{i}" for i in num_cols]
            w.writerow(hdr)
        for row in matrix[start_data:]:
            if not row:
                continue
            name = row[0] if row else ""
            values = [row[i] if i < len(row) else "" for i in num_cols]
            if not name and not any(v for v in values):
                continue
            # Filter: keep rows that have any digit in values
            if not any(any(ch.isdigit() for ch in (v or "")) for v in values):
                continue
            w.writerow([name] + values)
            written += 1
    return written


class OfficeManagerRunner:
    def __init__(self, driver: webdriver.Chrome, csv_path: Path, wait_timeout: int = 25) -> None:
        self.driver = driver
        self.csv_path = csv_path
        self.wait = WebDriverWait(driver, wait_timeout)

    # ---------- Navigation / auth ----------
    def ensure_role_selected(self) -> None:
        if "/SelectRole" in self.driver.current_url:
            # Wait roles to appear
            try:
                self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '[name="roleId"]')))
            except Exception:
                return
            # Log available roles
            try:
                roles = self.driver.execute_script(
                    "return Array.from(document.querySelectorAll('[name=\"roleId\"]')).map(e=>({val:e.value||e.getAttribute('value')||'', text:(e.textContent||e.value||'').trim()}));"
                ) or []
                if roles:
                    print("[role] Доступные роли:")
                    for r in roles:
                        print(f"[role] value={r.get('val')} text={r.get('text')}")
            except Exception:
                pass
            # Try click by exact value
            clicked = False
            for sel in (f'button[name="roleId"][value="{ROLE_ID}"]', f'[name="roleId"][value="{ROLE_ID}"]'):
                try:
                    el = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    except Exception:
                        pass
                    el.click()
                    clicked = True
                    break
                except Exception:
                    continue
            # Fallback: click the first role if specific value not found (to move forward)
            if not clicked:
                try:
                    first = self.driver.find_element(By.CSS_SELECTOR, '[name="roleId"]')
                    first.click()
                    clicked = True
                except Exception:
                    pass
            # Wait to leave SelectRole
            if clicked:
                try:
                    WebDriverWait(self.driver, 10).until(lambda d: "/SelectRole" not in d.current_url)
                except Exception:
                    print("[role] Не удалось покинуть SelectRole автоматически. Проверьте ROLE_ID.")

    def open_select_department(self) -> None:
        self.driver.get(SELECT_DEPARTMENT_URL)
        self.ensure_role_selected()
        if "/SelectDepartment" not in self.driver.current_url:
            try:
                self.driver.get(SELECT_DEPARTMENT_URL)
            except Exception:
                pass

    def get_cities(self):
        self.open_select_department()
        print(f"[nav] Текущий URL: {self.driver.current_url}")
        # If still on SelectRole, try once more to select the role
        if "/SelectRole" in self.driver.current_url:
            self.ensure_role_selected()
            self.open_select_department()
        self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '*[name="uuid"]')))
        items = self.driver.execute_script(
            """
            return Array.from(document.querySelectorAll('*[name="uuid"]')).map(b=>({
              name:(b.textContent||'').trim(),
              uuid:b.getAttribute('value')||b.getAttribute('data-value')||b.getAttribute('data-uuid')||b.getAttribute('uuid')||b.getAttribute('data-id')||''
            })).filter(x=>x.name && x.uuid);
            """
        ) or []
        seen=set(); cities=[]
        for it in items:
            uuid=it.get('uuid'); name=it.get('name')
            if uuid and name and uuid not in seen:
                seen.add(uuid); cities.append((name, uuid))
        cities.sort(key=lambda x: x[0].lower())
        return cities

    def select_city(self, city_uuid: str) -> None:
        self.open_select_department()
        self.ensure_role_selected()
        clicked = False
        # Try via JS matching any element with name=uuid and matching id across attributes
        try:
            clicked = bool(
                self.driver.execute_script(
                    """
                    var uuid = arguments[0];
                    var nodes = Array.from(document.querySelectorAll('*[name="uuid"]'));
                    var el = nodes.find(function(n){
                      var v = n.getAttribute('value') || n.getAttribute('data-value') || n.getAttribute('data-uuid') || n.getAttribute('uuid') || n.getAttribute('data-id') || '';
                      return v === uuid;
                    });
                    if(el){ try{ el.scrollIntoView({block:'center'}); }catch(e){}
                      el.click(); return true; }
                    return false;
                    """,
                    city_uuid,
                )
            )
        except Exception:
            clicked = False

        # Fallback to specific selectors
        if not clicked:
            for sel in (
                f'button[name="uuid"][value="{city_uuid}"]',
                f'a[name="uuid"][value="{city_uuid}"]',
                f'*[name="uuid"][value="{city_uuid}"]',
            ):
                try:
                    el = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    except Exception:
                        pass
                    el.click()
                    clicked = True
                    break
                except Exception:
                    continue

        # Wait a moment and ensure navigation proceeds
        if clicked:
            for _ in range(40):
                try:
                    if "/SelectDepartment" not in self.driver.current_url:
                        break
                except Exception:
                    pass
                time.sleep(0.05)

    # ---------- Report helpers ----------
    def open_report(self) -> None:
        self.driver.get(REPORT_URL)
        self.ensure_role_selected()
        # Wait for department multiselect to become available
        try:
            self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#SelectedUnitIds option')))
        except Exception:
            time.sleep(0.2)

    def compute_dates(self):
        today = __import__('datetime').date.today()
        start = today.replace(day=1)
        yesterday = today - __import__('datetime').timedelta(days=1)
        return start, yesterday

    def set_period(self, start_date, end_date):
        s = start_date.strftime("%d.%m.%Y"); e = end_date.strftime("%d.%m.%Y")
        js = """
        function setVal(sel, val){
          var el=document.querySelector(sel); if(!el) return false;
          el.value=val;
          try{ el.dispatchEvent(new Event('input',{bubbles:true})); }catch(e){ var ev=document.createEvent('HTMLEvents'); ev.initEvent('input',true,false); el.dispatchEvent(ev); }
          try{ el.dispatchEvent(new Event('change',{bubbles:true})); }catch(e){ var ev2=document.createEvent('HTMLEvents'); ev2.initEvent('change',true,false); el.dispatchEvent(ev2); }
          return true;
        }
        return [setVal('#StartDate', arguments[0]), setVal('#EndDate', arguments[1])];
        """
        try:
            self.driver.execute_script(js, s, e)
        except Exception:
            pass

    def get_departments(self):
        names = []
        for _ in range(100):
            try:
                names = self.driver.execute_script(
                    "return Array.from(document.querySelectorAll('#SelectedUnitIds option')).map(o=>(o.text||'').trim()).filter(Boolean);"
                ) or []
            except Exception:
                names = []
            if names:
                break
            time.sleep(0.1)
        return names

    def select_only_department(self, dept_name: str):
        try:
            self.driver.execute_script(
                """
                var s=document.getElementById('SelectedUnitIds'); if(!s) return false;
                var name=arguments[0];
                Array.from(s.options).forEach(o => o.selected=((o.text||'').trim()===name));
                var e; try{e=new Event('change',{bubbles:true});}catch(err){e=document.createEvent('HTMLEvents'); e.initEvent('change',true,false);} s.dispatchEvent(e);
                return true;
                """,
                dept_name,
            )
        except Exception:
            pass

    def click_build(self):
        try:
            btn = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#buildReportButton, [name="reportButton"]')))
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            except Exception:
                pass
            btn.click()
        except Exception:
            try:
                self.driver.execute_script("if(window.buildReport){buildReport();}")
            except Exception:
                pass

    def append_csv_row(self, row):
        with self.csv_path.open("a", newline="", encoding="utf-8-sig") as f:
            import csv as _csv
            _csv.writer(f, delimiter=";").writerow(row)

    def reset_csv(self):
        with self.csv_path.open("w", newline="", encoding="utf-8-sig") as f:
            f.write("\ufeff")

    def run(self) -> int:
        self.reset_csv()
        cities = self.get_cities()
        print(f"[cities] Найдено: {len(cities)} — {', '.join([c[0] for c in cities])}")
        start, end = self.compute_dates()
        print(f"[dates] Период: {start:%d.%m.%Y} — {end:%d.%m.%Y}")

        for cidx, (city_name, city_uuid) in enumerate(cities, start=1):
            print("\n" + "#" * 80)
            print(f"[city] ({cidx}/{len(cities)}) {city_name}")
            try:
                self.select_city(city_uuid)
                self.open_report()
                self.set_period(start, end)
                departments = self.get_departments()
                print(f"[depts] {departments}")
                self.append_csv_row([f"ГОРОД: {city_name}"])

                for didx, dept in enumerate(departments, start=1):
                    print("\n" + "=" * 80)
                    print(f"[dept] ({didx}/{len(departments)}) {dept}")
                    self.append_csv_row([f"ОТДЕЛ: {dept}"])
                    self.select_only_department(dept)

                    old_html = None
                    try:
                        old_html = self.driver.find_element(By.CSS_SELECTOR, "#report").get_attribute("innerHTML")
                    except Exception:
                        pass
                    self.click_build()
                    if old_html is not None:
                        for _ in range(200):
                            try:
                                if self.driver.find_element(By.CSS_SELECTOR, "#report").get_attribute("innerHTML") != old_html:
                                    break
                            except Exception:
                                pass
                            time.sleep(0.05)

                    # Dump current table
                    try:
                        _ = dump_table_to_csv(self.driver, self.csv_path)
                    except Exception as e:
                        print(f"[csv] dump failed: {e}")
                    # separator
                    self.append_csv_row([""])

            except Exception as e:
                print(f"[warn] Ошибка по городу {city_name}: {e}")
                self.append_csv_row([f"ГОРОД: {city_name}", f"ОШИБКА: {e}"])
                self.append_csv_row([""])

            # Return to role selection between cities to reset context
            try:
                self.back_to_select_role()
            except Exception:
                pass

        print(f"[done] CSV: {self.csv_path}")
        return 0

    def back_to_select_role(self) -> None:
        try:
            self.driver.get(BACK_TO_SELECT_ROLE_URL)
        except Exception:
            pass
        try:
            WebDriverWait(self.driver, 10).until(EC.url_contains("/SelectRole"))
        except Exception:
            pass
        self.ensure_role_selected()
        self.open_select_department()


if __name__ == "__main__":
    raise SystemExit(main())
