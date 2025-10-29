import os
import sys
import csv
import time
import json
import subprocess
import urllib.parse
import datetime as dt
from pathlib import Path
from glob import glob
from typing import List, Optional, Tuple, Dict, Any

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

try:
    from webdriver_manager.chrome import ChromeDriverManager  # type: ignore
except Exception:  # pragma: no cover
    ChromeDriverManager = None  # type: ignore


# =========================
# Defaults pinned in script
# =========================

# Port is relevant for remote-debugging flows on Windows; kept for reference.
PORT = 9222

# Pinned CSV file path (store reports under ./reports on host)
CSV_FILE = "reports/MaterialWriteOffReport.csv"

# Pinned URLs and role for Project Manager scenario
REPORT_URL = "https://officemanager.dodopizza.ru/OfficeManager/Debiting/PrepareExcelReport"
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ROLE_ID = "8"

# Optional slow delay can still be overridden via env
SLOW_DELAY = float(os.environ.get("SLOW_DELAY", "0"))
STEALTH = True  # включено по умолчанию; можно отключить переменной среды

# Supabase config defaults
SUPABASE_CONFIG_FILE = Path(os.environ.get("SUPABASE_CONFIG_FILE", "config/api"))
SUPABASE_KEY_FILE = Path(os.environ.get("SUPABASE_KEY_FILE", "config/api.key"))
BUGREPORT_FILE = Path("reports/bugrepot.txt")


def log_bugreport(context: str, count: int, stdout: str, stderr: str) -> None:
    try:
        BUGREPORT_FILE.parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    try:
        timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with BUGREPORT_FILE.open("a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {context} — {count} rows\n")
            if stdout:
                f.write(stdout.strip() + "\n")
            if stderr:
                f.write(f"stderr: {stderr.strip()}\n")
            f.write("\n")
    except Exception as exc:
        print(f"[BUGLOG] Не удалось записать ответ БД: {exc}")


# =========================
# Helpers / driver bootstrap
# =========================

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

    # Reuse existing authenticated profile
    options.add_argument(f"--user-data-dir={str(user_data_dir)}")
    (user_data_dir / "Default").mkdir(parents=True, exist_ok=True)

    # Best-effort: remove stale lock files from a mounted profile
    cleanup_profile_locks(user_data_dir)

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")

    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    if STEALTH:
        options.add_experimental_option("excludeSwitches", ["enable-automation"])  # hide automation banner
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--lang=ru-RU")

    chrome_bin = os.environ.get("CHROME_BIN")
    if chrome_bin:
        options.binary_location = chrome_bin

    driver_path = os.environ.get("CHROMEDRIVER")
    if driver_path and Path(driver_path).exists():
        service = Service(executable_path=driver_path)
    else:
        if ChromeDriverManager is None:
            raise RuntimeError(
                "Chromedriver not found and webdriver-manager is unavailable."
            )
        service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=options)
    if STEALTH:
        try:
            _apply_stealth(driver)
        except Exception:
            pass
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


# =========================
# Project Manager runner
# =========================

class ProjectManagerRunner:
    def __init__(
        self,
        driver: webdriver.Chrome,
        role_id: str,
        select_department_url: str,
        back_to_select_role_url: str,
        report_url: str,
        csv_file: Path,
        wait_timeout: int = 25,
        slow_delay: float = 0.0,
    ) -> None:
        self.driver = driver
        self.role_id = role_id
        self.select_department_url = select_department_url
        self.back_to_select_role_url = back_to_select_role_url
        self.report_url = report_url
        self.csv_file = csv_file
        self.wait = WebDriverWait(self.driver, wait_timeout)
        self.slow_delay = slow_delay
        config_url, config_key = self._read_supabase_config()
        project_id = os.environ.get("SUPABASE_PROJECT_ID", "").strip()
        self.supabase_url = (os.environ.get("SUPABASE_URL") or "").strip()
        if not self.supabase_url and config_url:
            self.supabase_url = config_url
        if not self.supabase_url and project_id:
            self.supabase_url = f"https://{project_id}.supabase.co"
        self.supabase_key = (os.environ.get("SUPABASE_SERVICE_KEY") or os.environ.get("SUPABASE_KEY") or "").strip()
        if not self.supabase_key:
            self.supabase_key = config_key
        if not self.supabase_key:
            self.supabase_key = self._read_supabase_key_file()
        self.supabase_table = os.environ.get("SUPABASE_TABLE", "material_write_offs").strip() or "material_write_offs"
        self.supabase_on_conflict = os.environ.get("SUPABASE_ON_CONFLICT", "city,department,dt").strip()
        self.supabase_batch_size = max(1, int(os.environ.get("SUPABASE_BATCH_SIZE", "50")))
        self.supabase_timeout = float(os.environ.get("SUPABASE_TIMEOUT", "15"))
        self.supabase_enabled = bool(self.supabase_url and self.supabase_key)
        self._supabase_warned = False
        self._supabase_conflict_forced_off = False
        self._supabase_conflict_notice_shown = False
        default_field_map = {
            "city": os.environ.get("SUPABASE_FIELD_CITY", "city"),
            "department": os.environ.get("SUPABASE_FIELD_DEPARTMENT", "department"),
            "date": os.environ.get("SUPABASE_FIELD_DATE", "dt"),
            "value": os.environ.get("SUPABASE_FIELD_VALUE", "write_off_amount"),
        }
        field_map_override = os.environ.get("SUPABASE_FIELD_MAP", "").strip()
        if field_map_override:
            try:
                parsed_map = json.loads(field_map_override)
                if isinstance(parsed_map, dict):
                    default_field_map.update({k: str(v) for k, v in parsed_map.items() if v})
            except json.JSONDecodeError:
                print("[SUPABASE] Не удалось разобрать SUPABASE_FIELD_MAP как JSON — использую значения по умолчанию.")
        self.supabase_field_map = default_field_map

    # ---------- Navigation / auth ----------
    def ensure_role_selected(self) -> None:
        if "/SelectRole" in self.driver.current_url:
            # Log available roles to help choose ROLE_ID
            try:
                roles = self.driver.execute_script(
                    """
                    return Array.from(document.querySelectorAll('[name="roleId"]'))
                      .map(el => ({
                        tag: el.tagName,
                        type: el.getAttribute('type') || '',
                        value: el.getAttribute('value') || '',
                        text: (el.textContent||el.value||'').trim()
                      }));
                    """
                ) or []
                if roles:
                    print("[role] Доступные роли:")
                    for r in roles:
                        print(f"[role] value={r.get('value')} text={r.get('text')}")
            except Exception:
                pass
            try:
                self.wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, f'button[name="roleId"][value="{self.role_id}"]')
                    )
                ).click()
            except Exception:
                try:
                    self.wait.until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, f'[name="roleId"][value="{self.role_id}"]')
                        )
                    ).click()
                except Exception:
                    pass
            try:
                WebDriverWait(self.driver, 10).until(
                    lambda d: "/SelectRole" not in d.current_url
                )
            except Exception:
                pass

    def open_select_department(self) -> None:
        for attempt in range(2):
            try:
                self.driver.get(self.select_department_url)
            except Exception:
                pass
            self.ensure_role_selected()
            if "/SelectDepartment" in (self.driver.current_url or ""):
                return
            if attempt == 0:
                # Попробуем сбросить привязку отдела и вернуться к SelectRole
                try:
                    self.driver.get(self.back_to_select_role_url)
                except Exception:
                    pass
                try:
                    WebDriverWait(self.driver, 10).until(EC.url_contains("/SelectRole"))
                except Exception:
                    pass
                self.ensure_role_selected()
        print(f"[nav] Не удалось напрямую открыть SelectDepartment, текущий URL: {self.driver.current_url}")

    def back_to_select_role(self) -> None:
        try:
            self.driver.get(self.back_to_select_role_url)
        except Exception:
            pass
        try:
            WebDriverWait(self.driver, 10).until(EC.url_contains("/SelectRole"))
        except Exception:
            pass
        self.ensure_role_selected()
        self.open_select_department()

    # ---------- City/department helpers ----------
    def get_cities(self) -> List[Tuple[str, str]]:
        self.open_select_department()
        print(f"[nav] Текущий URL: {self.driver.current_url}")
        # Wait for city selectors; support multiple tag types
        try:
            self.wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, '*[name="uuid"]'))
            )
        except Exception:
            # Extra hint if stuck on role selection
            if "/SelectRole" in self.driver.current_url:
                print("[hint] Похоже, вы на странице выбора роли. Проверьте ROLE_ID.")
            raise
        try:
            items = self.driver.execute_script(
                """
                return Array.from(document.querySelectorAll('*[name="uuid"]'))
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
            raise RuntimeError("Не удалось получить список городов на SelectDepartment")
        return cities

    def select_city(self, city_uuid: str) -> None:
        self.open_select_department()
        self.ensure_role_selected()
        self.wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="uuid"][value="{city_uuid}"]'))
        ).click()
        time.sleep(0.2)

    def get_departments(self, limit: Optional[int] = None) -> List[str]:
        names: List[str] = []
        # Try to read from real <select id="UnitId">
        for _ in range(100):
            try:
                names = self.driver.execute_script(
                    "return Array.from(document.querySelectorAll('#UnitId option'))\n"
                    "  .map(o => (o.text||'').trim())\n"
                    "  .filter(t => t && t.toLowerCase()!=='выбрать все');"
                ) or []
            except Exception:
                names = []
            if names:
                break
            time.sleep(0.1)
        if not names:
            # Fallback: attempt to open the dropdown and read items
            try:
                opened = self.driver.execute_script(
                    "var s=document.getElementById('UnitId'); if(!s) return false;\n"
                    "var box=s.closest('.select-report'); if(!box) return false;\n"
                    "var cap=box.querySelector('.CaptionCont'); if(!cap) return false; cap.click(); return true;"
                )
                if opened:
                    time.sleep(0.3)
                names = self.driver.execute_script(
                    "return Array.from(document.querySelectorAll('.open li'))\n"
                    "  .map(li => (li.textContent||'').trim())\n"
                    "  .filter(t => t && t.toLowerCase()!=='выбрать все');"
                ) or []
            except Exception:
                names = []
        if limit is not None:
            names = names[: max(0, int(limit))]
        return names

    def force_select_only_one_by_text(self, dept_name: str) -> List[str]:
        try:
            selected = self.driver.execute_script(
                """
                var s=document.getElementById('UnitId');
                if(!s) return [];
                var name=arguments[0];
                Array.from(s.options).forEach(o => o.selected=((o.text||'').trim()===name));
                var e; try{e=new Event('change',{bubbles:true});}catch(err){e=document.createEvent('HTMLEvents'); e.initEvent('change',true,false);} s.dispatchEvent(e);
                return Array.from(s.selectedOptions).map(o=>(o.text||'').trim());
                """,
                dept_name,
            )
            if isinstance(selected, list):
                return [str(x) for x in selected]
            return []
        except Exception:
            return []

    def choose_department(self, dept_name: str) -> None:
        # Try up to 3 times to enforce a single selection
        for _ in range(3):
            chosen = self.force_select_only_one_by_text(dept_name)
            if len(chosen) == 1 and chosen[0] == dept_name:
                break
            time.sleep(0.1)
        # Align UI wrapper state (optional best-effort)
        try:
            self.driver.execute_script(
                """
                var s=document.getElementById('UnitId'); if(!s) return;
                var name=arguments[0];
                var box=s.closest('.select-report'); if(!box) return;
                var items=box.querySelectorAll('li');
                items.forEach(li=>{
                  var t=(li.textContent||'').trim();
                  var sel=li.classList.contains('selected');
                  if(t===name && !sel){ li.click(); }
                  if(t!==name && sel){ li.click(); }
                });
                """,
                dept_name,
            )
        except Exception:
            pass

    # ---------- Filters / report ----------
    def compute_dates(self) -> List[dt.date]:
        today = dt.date.today()
        start = today.replace(day=1)
        yesterday = today - dt.timedelta(days=1)
        if yesterday < start:
            return []
        days = (yesterday - start).days + 1
        return [start + dt.timedelta(days=i) for i in range(days)]

    def _parse_date(self, s: str) -> Optional[dt.date]:
        s = (s or "").strip()
        today = dt.date.today()
        if not s:
            return today - dt.timedelta(days=1)
        if s.isdigit():
            if len(s) <= 2:
                try:
                    return dt.date(today.year, today.month, int(s))
                except Exception:
                    return None
            if len(s) == 8:
                try:
                    return dt.datetime.strptime(s, "%d%m%Y").date()
                except Exception:
                    return None
        for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(s, fmt).date()
            except Exception:
                continue
        return None

    def prompt_dates(self) -> List[dt.date]:
        while True:
            try:
                start_raw = input("[INPUT] Начальная дата (ДД.ММ.ГГГГ): ").strip()
            except EOFError:
                start_raw = ""
            start = self._parse_date(start_raw)
            if not start:
                print("[INPUT] Некорректная дата. Пример: 01.09.2025")
                continue

            try:
                end_raw = input("[INPUT] Конечная дата (ДД.ММ.ГГГГ): ").strip()
            except EOFError:
                end_raw = ""
            end = self._parse_date(end_raw)
            if not end:
                print("[INPUT] Некорректная дата. Пример: 30.09.2025")
                continue

            if end < start:
                print("[INPUT] Конечная дата раньше начальной — поменял местами.")
                start, end = end, start
            days = (end - start).days + 1
            return [start + dt.timedelta(days=i) for i in range(days)]

    def set_period_dates(self, d: dt.date) -> None:
        # For Project Manager: set StartDate/EndDate directly and dispatch events
        start_s = d.strftime("%d.%m.%Y")
        end_s = d.strftime("%d.%m.%Y")
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
            self.driver.execute_script(js, start_s, end_s)
        except Exception:
            pass

    def open_report(self) -> None:
        self.driver.get(self.report_url)
        self.ensure_role_selected()
        # Some pages require a short delay for scripts to wire up
        time.sleep(0.2)

    def select_all_reasons(self) -> None:
        # Select all options in DebitingReasonId if present (stabilizes totals)
        try:
            self.driver.execute_script(
                "var s=document.getElementById('DebitingReasonId'); if(!s) return;"
                "Array.from(s.options).forEach(o=>o.selected=true);"
                "try{s.dispatchEvent(new Event('change',{bubbles:true}));}catch(e){var ev=document.createEvent('HTMLEvents');ev.initEvent('change',true,false);s.dispatchEvent(ev);}"
            )
        except Exception:
            pass

    def click_build_report(self) -> None:
        try:
            btn = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '[name="reportButton"], #buildReportButton'))
            )
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            except Exception:
                pass
            btn.click()
        except Exception:
            # Fallback: try JS entry points if exist
            try:
                self.driver.execute_script("if(window.buildReport){buildReport();}")
            except Exception:
                pass

    def read_total_value(self) -> str:
        # Prefer explicit total cells, then fallback to last numeric cell
        selectors = [
            "tbody td.totalValue",
            "tfoot td",
            "tbody tr:last-child td:last-child",
            "tbody td",
        ]
        for sel in selectors:
            try:
                elems = self.driver.find_elements(By.CSS_SELECTOR, sel)
            except Exception:
                elems = []
            if not elems:
                continue
            candidates = [e for e in elems if (e.text or "").strip() and any(c.isdigit() for c in e.text)]
            if not candidates:
                continue
            target = candidates[-1]
            try:
                self.driver.execute_script(
                    "arguments[0].style.backgroundColor='#00ff00';arguments[0].style.color='#000';",
                    target,
                )
            except Exception:
                pass
            txt = (target.text or "").replace("\xa0", " ").strip()
            # Normalize Russian currency formatting (spaces as thousands, comma or dot as decimal)
            txt = txt.replace("₽", "").replace(" ", "")
            return txt
        return ""

    # ---------- Supabase helpers ----------
    def _read_supabase_config(self) -> Tuple[str, str]:
        try:
            if not SUPABASE_CONFIG_FILE.is_file():
                return "", ""
            raw = SUPABASE_CONFIG_FILE.read_text(encoding="utf-8").strip()
            if not raw:
                return "", ""
            data = {}
            if raw.lstrip().startswith("{"):
                data = json.loads(raw)
                if not isinstance(data, dict):
                    print(f"[SUPABASE] Ожидал объект JSON в {SUPABASE_CONFIG_FILE}")
                    return "", ""
            else:
                for line in raw.splitlines():
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" in line:
                        key, value = line.split("=", 1)
                        data[key.strip().lower()] = value.strip()
            url = str(data.get("url", "")).strip()
            key = ""
            for candidate in ("key", "apikey", "anon_key", "service_key"):
                if candidate in data and data[candidate]:
                    key = str(data[candidate]).strip()
                    break
            if not key:
                for candidate_name, candidate_value in data.items():
                    if "key" in candidate_name and candidate_value:
                        key = str(candidate_value).strip()
                        break
            return url, key
        except Exception as exc:
            print(f"[SUPABASE] Не удалось прочитать {SUPABASE_CONFIG_FILE}: {exc}")
            return "", ""

    def _read_supabase_key_file(self) -> str:
        try:
            if SUPABASE_KEY_FILE.is_file():
                value = SUPABASE_KEY_FILE.read_text(encoding="utf-8").strip()
                if value:
                    return value
        except Exception as exc:
            print(f"[SUPABASE] Не удалось прочитать ключ из {SUPABASE_KEY_FILE}: {exc}")
        return ""

    def _push_supabase_rows(self, rows: List[List[str]]) -> List[Dict[str, Any]]:
        if not self.supabase_enabled:
            if not self._supabase_warned:
                print("[SUPABASE] URL или ключ не заданы — пропускаю отправку.")
                self._supabase_warned = True
            return []
        payload: List[dict] = []
        for row in rows:
            record = self._build_supabase_record(row)
            if record is not None:
                payload.append(record)
        if not payload:
            return []
        responses: List[Dict[str, Any]] = []
        for chunk_start in range(0, len(payload), self.supabase_batch_size):
            chunk = payload[chunk_start:chunk_start + self.supabase_batch_size]
            responses.append(self._post_supabase_chunk(chunk))
        return responses

    def _build_supabase_record(self, row: List[str]) -> Optional[dict]:
        if len(row) < 4:
            return None
        city, dept, date_str, value = row[:4]
        fm = self.supabase_field_map
        return {
            fm["city"]: city,
            fm["department"]: dept,
            fm["date"]: self._safe_parse_date(date_str),
            fm["value"]: self._parse_number(value),
        }

    def _parse_number(self, value: str) -> Optional[float]:
        txt = (value or "").strip().replace("\xa0", "").replace(" ", "")
        if not txt:
            return None
        txt = txt.replace(",", ".")
        try:
            number = float(txt)
            if number.is_integer():
                return int(number)
            return number
        except ValueError:
            return None

    def _safe_parse_date(self, value: str) -> str:
        txt = (value or "").strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(txt, fmt).date().isoformat()
            except Exception:
                continue
        return txt

    def _post_supabase_chunk(self, chunk: List[dict], allow_conflict: bool = True) -> Dict[str, Any]:
        if not chunk:
            return {
                "count": 0,
                "stdout": "",
                "stderr": "",
                "endpoint": "",
                "http_status": 0,
                "on_conflict": False,
            }
        use_conflict = (
            allow_conflict
            and bool(self.supabase_on_conflict)
            and not self._supabase_conflict_forced_off
        )
        base_url = self.supabase_url.rstrip("/")
        endpoint = f"{base_url}/rest/v1/{self.supabase_table}"
        params: Dict[str, str] = {}
        if use_conflict:
            params["on_conflict"] = self.supabase_on_conflict
        if params:
            endpoint = f"{endpoint}?{urllib.parse.urlencode(params)}"
        prefer_parts = ["return=representation"]
        if use_conflict:
            prefer_parts.append("resolution=merge-duplicates")
        data = json.dumps(chunk, ensure_ascii=False)
        cmd = ["curl", "-sS"]
        if self.supabase_timeout > 0:
            cmd.extend(["--max-time", str(self.supabase_timeout)])
        cmd.extend(
            [
                "-w",
                "HTTPSTATUS:%{http_code}",
                "-X",
                "POST",
                endpoint,
                "-H",
                f"apikey: {self.supabase_key}",
                "-H",
                f"Authorization: Bearer {self.supabase_key}",
                "-H",
                "Content-Type: application/json",
                "-H",
                f"Prefer: {','.join(dict.fromkeys(prefer_parts))}",
                "-d",
                data,
            ]
        )
        result = subprocess.run(cmd, capture_output=True, text=True)
        stdout_raw = result.stdout or ""
        stderr_txt = (result.stderr or "").strip()
        http_status = 0
        body_txt = stdout_raw
        if "HTTPSTATUS:" in stdout_raw:
            body_txt, _, status_part = stdout_raw.rpartition("HTTPSTATUS:")
            try:
                http_status = int((status_part or "").strip() or "0")
            except ValueError:
                http_status = 0
        body_txt = body_txt.strip()
        if result.returncode != 0 and http_status == 0:
            raise RuntimeError(f"curl exit code {result.returncode}: {stderr_txt or stdout_raw}")
        if http_status >= 400:
            error_payload = body_txt or stderr_txt
            conflict_missing = False
            if use_conflict:
                lowered = (error_payload or "").lower()
                conflict_missing = "no unique or exclusion constraint" in lowered
                if not conflict_missing:
                    try:
                        parsed = json.loads(body_txt or "{}")
                        message = str(parsed.get("message", "")).lower()
                        conflict_missing = "no unique or exclusion constraint" in message
                    except Exception:
                        pass
            if use_conflict and conflict_missing:
                if not self._supabase_conflict_notice_shown:
                    print("[SUPABASE] У таблицы нет уникального индекса для on_conflict — повторяю без него.")
                    self._supabase_conflict_notice_shown = True
                self._supabase_conflict_forced_off = True
                retry_info = self._post_supabase_chunk(chunk, allow_conflict=False)
                extra_msg = "[SUPABASE] Повтор без on_conflict из-за отсутствия уникального индекса."
                prev_stdout = (retry_info.get("stdout") or "").strip()
                retry_info["stdout"] = (
                    f"{prev_stdout}\n{extra_msg}".strip() if prev_stdout else extra_msg
                )
                return retry_info
            raise RuntimeError(f"Supabase HTTP {http_status}: {error_payload or 'unknown error'}")
        print(f"[SUPABASE] Записано {len(chunk)} записей.")
        return {
            "count": len(chunk),
            "stdout": body_txt,
            "stderr": stderr_txt,
            "endpoint": endpoint,
            "http_status": http_status,
            "on_conflict": use_conflict,
        }

    # ---------- CSV ----------
    def reset_csv(self) -> None:
        with open(self.csv_file, "w", encoding="utf-8-sig", newline="") as f:
            f.write("\ufeff")

    def append_csv_row(self, row: List[str]) -> None:
        with open(self.csv_file, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter=";").writerow(row)

    def append_csv_rows(self, rows: List[List[str]]) -> None:
        if not rows:
            return
        with open(self.csv_file, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter=";").writerows(rows)

    # ---------- Main flow ----------
    def run(self) -> int:
        self.open_select_department()
        cities = self.get_cities()
        print(f"[CITIES] Найдено: {len(cities)} — {', '.join([c[0] for c in cities])}")

        dates = self.prompt_dates()
        if dates:
            print(f"[DATES] Диапазон: {dates[0]:%d.%m.%Y} — {dates[-1]:%d.%m.%Y} (всего {len(dates)})")
        else:
            print("[DATES] Диапазон дат пуст — ничего обрабатывать.")

        self.reset_csv()

        for cidx, (city_name, city_uuid) in enumerate(cities, start=1):
            print("\n" + "#" * 80)
            print(f"[CITY] ({cidx}/{len(cities)}) {city_name}")
            try:
                self.select_city(city_uuid)
                self.open_report()
                self.select_all_reasons()

                # Departments for this city
                departments = self.get_departments(limit=None)
                print(f"[DEPTS] {departments}")
                city_csv_rows: List[List[str]] = [[f"ГОРОД: {city_name}", ""]]
                city_sup_rows: List[List[str]] = []

                for didx, dept in enumerate(departments, start=1):
                    print("\n" + "=" * 80)
                    print(f"[DEPT] ({didx}/{len(departments)}) {dept}")
                    self.choose_department(dept)
                    city_csv_rows.append([f"ОТДЕЛ: {dept}", ""])

                    for d in dates:
                        self.set_period_dates(d)
                        old_html = None
                        try:
                            old_html = self.driver.find_element(By.CSS_SELECTOR, "#report").get_attribute(
                                "innerHTML"
                            )
                        except Exception:
                            pass
                        self.click_build_report()
                        if old_html is not None:
                            for _ in range(200):
                                try:
                                    if (
                                        self.driver.find_element(By.CSS_SELECTOR, "#report").get_attribute(
                                            "innerHTML"
                                        )
                                        != old_html
                                    ):
                                        break
                                except Exception:
                                    pass
                                time.sleep(0.05)
                        val = self.read_total_value()
                        city_csv_rows.append([d.strftime("%d.%m.%Y"), val])
                        city_sup_rows.append([city_name, dept, d.strftime("%d.%m.%Y"), val])
                        print(f"[CSV] {d:%d.%m.%Y}: {val}")

                if city_csv_rows:
                    self.append_csv_rows(city_csv_rows)
                if city_sup_rows:
                    try:
                        responses = self._push_supabase_rows(city_sup_rows)
                        if responses:
                            for idx, resp in enumerate(responses, start=1):
                                context = (
                                    f"MaterialWriteOffReport — {city_name} chunk {idx}/{len(responses)} "
                                    f"(rows: {len(city_sup_rows)})"
                                )
                                log_bugreport(
                                    context,
                                    int(resp.get("count", 0)),
                                    str(resp.get("stdout", "")),
                                    str(resp.get("stderr", "")),
                                )
                        print(f"[SUPABASE] {city_name}: выгружено {len(city_sup_rows)} записей.")
                    except Exception as exc:
                        error_text = f"Ошибка отправки: {exc}"
                        print(f"[SUPABASE] {error_text}")
                        log_bugreport(
                            f"MaterialWriteOffReport — ERROR ({city_name}, rows: {len(city_sup_rows)})",
                            0,
                            "",
                            error_text,
                        )

            except Exception as e:
                print(f"[WARN] Ошибка при обработке города {city_name}: {e}")
                self.append_csv_row([f"ГОРОД: {city_name}", f"ОШИБКА: {e}"])
                self.append_csv_row(["", ""])  # separator

            # Return to SelectRole between cities
            try:
                self.back_to_select_role()
            except Exception:
                pass

        print(f"[DONE] Готово! Файл {self.csv_file} сохранён.")
        return 0


def main() -> int:
    # Env/config
    user_data_dir = Path(os.environ.get("USER_DATA_DIR") or (Path.cwd() / "profile"))
    user_data_dir.mkdir(parents=True, exist_ok=True)

    headless = env_bool("HEADLESS", True)

    # Use pinned defaults (no need to set env before running)
    role_id = ROLE_ID
    select_department_url = SELECT_DEPARTMENT_URL
    back_to_select_role_url = BACK_TO_SELECT_ROLE_URL
    report_url = REPORT_URL
    csv_file = Path(CSV_FILE)
    # Ensure reports directory exists
    try:
        csv_file.parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    slow_delay = SLOW_DELAY

    print(
        f"[run] Chrome headless={headless}; profile={user_data_dir}; report_url={report_url}",
        flush=True,
    )

    try:
        driver = build_chrome(headless=headless, user_data_dir=user_data_dir)
    except Exception as e:
        print(f"[run] Failed to launch Chrome: {e}", file=sys.stderr)
        return 2

    try:
        runner = ProjectManagerRunner(
            driver=driver,
            role_id=role_id,
            select_department_url=select_department_url,
            back_to_select_role_url=back_to_select_role_url,
            report_url=report_url,
            csv_file=csv_file,
            wait_timeout=25,
            slow_delay=slow_delay,
        )
        return runner.run()
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
