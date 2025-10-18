from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from typing import Optional, List, Tuple, Dict, Any
from pathlib import Path

import csv
import datetime
import os
import socket
import subprocess
import time
import itertools
import json
import urllib.parse


PORT = 9223
CSV_FILE = "reports/LossesAndExcessReport.csv"
BUGREPORT_FILE = Path("reports/bugrepot.txt")
REPORT_URL = "https://officemanager.dodopizza.ru/InventoryManager/LossesAndExcees"
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ROLE_ID = os.environ.get("ROLE_ID", "7")
SLOW_DELAY = float(os.environ.get("SLOW_DELAY", "0"))
STEALTH = os.environ.get("STEALTH", "1")
SUPABASE_CONFIG_FILE = Path(os.environ.get("SUPABASE_CONFIG_FILE", "config/api"))
SUPABASE_KEY_FILE = Path(os.environ.get("SUPABASE_KEY_FILE", "config/api.key"))


class LossesAndExcessReporter:
    def __init__(self, port: int = PORT, csv_file: str = CSV_FILE, url: str = REPORT_URL, slow: float = SLOW_DELAY):
        self.port = port
        self.csv_file = csv_file
        self.url = url
        self.slow = slow
        self.driver: Optional[webdriver.Chrome] = None
        self.wait: Optional[WebDriverWait] = None
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
        self.supabase_table = os.environ.get("SUPABASE_TABLE", "losses_and_excess_reports").strip() or "losses_and_excess_reports"
        self.supabase_on_conflict = os.environ.get("SUPABASE_ON_CONFLICT", "city,department,dt,revisions").strip()
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
            "revisions": os.environ.get("SUPABASE_FIELD_REVISIONS", "revisions"),
            "metric": os.environ.get("SUPABASE_FIELD_METRIC", "metric_name"),
            "percent": os.environ.get("SUPABASE_FIELD_PERCENT", "percent_value"),
            "amount": os.environ.get("SUPABASE_FIELD_AMOUNT", "amount_value"),
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
        self._supabase_buffer: List[List[str]] = []

    def launch_chrome(self):
        print("[INIT] Настройка Chrome…")
        if os.name == 'nt':
            try:
                subprocess.run("taskkill /F /IM chrome.exe 2>nul", shell=True)
            except Exception:
                pass
            try:
                chrome_exe = rf"{os.environ.get('ProgramFiles','')}\\Google\\Chrome\\Application\\chrome.exe"
            except Exception:
                chrome_exe = None
            if chrome_exe and os.path.exists(chrome_exe):
                user_dir = os.environ.get("USER_DATA_DIR") or os.path.join(os.getcwd(), "profile")
                try:
                    os.makedirs(user_dir, exist_ok=True)
                except Exception:
                    pass
                subprocess.Popen([
                    chrome_exe,
                    f"--remote-debugging-port={self.port}",
                    f"--user-data-dir={user_dir}",
                    "--no-first-run",
                    "--no-default-browser-check",
                ])
                if not self._wait_port(self.port, 10):
                    raise RuntimeError(f"Порт {self.port} не открылся")
        else:
            print("[INIT] Linux/Docker: внешний Chrome не запускаю (использую драйвер).")

    def _make_service(self) -> Service:
        path = os.environ.get("CHROMEDRIVER", "/usr/bin/chromedriver")
        if path and os.path.exists(path):
            return Service(path)
        return Service(ChromeDriverManager().install())

    def connect_driver(self):
        print("[DRIVER] Инициализация драйвера Chrome…")
        options = webdriver.ChromeOptions()
        if STEALTH == "1":
            options.add_experimental_option("excludeSwitches", ["enable-automation"])  # скрыть баннер automation
            options.add_experimental_option("useAutomationExtension", False)
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--lang=ru-RU")
            options.add_argument("--window-size=1920,1080")
        if self._wait_port(self.port, 1):
            print("[DRIVER] Найден debuggerAddress — подключаюсь к внешнему Chrome…")
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{self.port}")
            self.driver = webdriver.Chrome(service=self._make_service(), options=options)
        else:
            if os.environ.get("CHROME_BIN"):
                options.binary_location = os.environ["CHROME_BIN"]
            user_dir = os.environ.get("USER_DATA_DIR") or os.path.join(os.getcwd(), "profile")
            try:
                os.makedirs(user_dir, exist_ok=True)
            except Exception:
                pass
            options.add_argument(f"--user-data-dir={user_dir}")
            if os.environ.get("HEADLESS", "1") == "1":
                options.add_argument("--headless=new")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--disable-gpu")
            self.driver = webdriver.Chrome(service=self._make_service(), options=options)
        self.wait = WebDriverWait(self.driver, 25)
        if STEALTH == "1":
            self._apply_stealth()

    def _apply_stealth(self):
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
                self.driver.execute_cdp_cmd("Network.setUserAgentOverride", {
                    "userAgent": new_ua or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
                    "platform": "Win32"
                })
            except Exception:
                pass
            try:
                self.driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {
                    "headers": {"Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"}
                })
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

    # ---------- Навигация и авторизация ----------
    def open_select_department(self):
        print("[NAV] Перехожу на экран выбора города…")
        self.driver.get(SELECT_DEPARTMENT_URL)
        self.ensure_role_selected()
        if "/SelectDepartment" not in self.driver.current_url:
            try:
                self.driver.get(SELECT_DEPARTMENT_URL)
            except Exception:
                pass

    def choose_role(self):
        print(f"[AUTH] Выбираю роль {ROLE_ID}…")
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

    def ensure_role_selected(self, city_uuid: Optional[str] = None):
        if "/SelectRole" in self.driver.current_url:
            self.choose_role()
            if city_uuid:
                try:
                    self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="uuid"][value="{city_uuid}"]'))).click()
                except Exception:
                    pass

    def get_cities(self) -> List[Tuple[str, str]]:
        print("[CITIES] Собираю список городов…")
        self.open_select_department()
        self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'button[name="uuid"], a[name="uuid"]')))
        try:
            items = self.driver.execute_script(
                """
                return Array.from(document.querySelectorAll('button[name="uuid"], a[name="uuid"]'))
                  .map(b => ({
                    name: (b.textContent || '').trim(),
                    uuid: b.getAttribute('value') || b.getAttribute('data-value') || b.getAttribute('data-uuid') ||
                          b.getAttribute('uuid') || b.getAttribute('data-id') || '',
                    tag: b.tagName
                  }))
                  .filter(x => x.name && x.uuid);
                """
            ) or []
        except Exception:
            items = []
        seen = set()
        cities: List[Tuple[str, str]] = []
        for it in items:
            uuid = it.get('uuid')
            name = it.get('name')
            if uuid and uuid not in seen:
                seen.add(uuid)
                cities.append((name, uuid))
        cities.sort(key=lambda x: x[0].lower())
        if not cities:
            raise RuntimeError("Не удалось получить список городов")
        print(f"[CITIES] Найдено городов: {len(cities)}")
        return cities

    def select_city(self, city_uuid: str):
        self.open_select_department()
        self.ensure_role_selected()
        try:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="uuid"][value="{city_uuid}"]'))).click()
        except Exception:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'a[name="uuid"][value="{city_uuid}"]'))).click()
        time.sleep(0.2)

    def open_report_for_city(self, city_uuid: str):
        print("[NAV] Перехожу на страницу LossesAndExcees…")
        self.driver.get(REPORT_URL)
        self.ensure_role_selected(city_uuid)

    def back_to_select_role(self):
        print("[NAV] Возврат на SelectRole…")
        try:
            self.driver.get(BACK_TO_SELECT_ROLE_URL)
        except Exception:
            pass
        try:
            WebDriverWait(self.driver, 10).until(EC.url_contains("/SelectRole"))
        except Exception:
            pass
        if "/SelectRole" in self.driver.current_url:
            self.choose_role()
        self.open_select_department()

    def refresh_report_page(self):
        print("[NAV] Обновляю страницу отчета…")
        try:
            self.driver.refresh()
            if self.wait:
                self.wait.until(EC.presence_of_element_located((By.ID, 'UnitId')))
        except Exception:
            time.sleep(1)

    # ---------- Элементы страницы LossesAndExcees ----------
    def get_units(self) -> List[Tuple[str, str]]:
        print("[UNITS] Получаю отделы (#UnitId)…")
        for _ in range(50):
            try:
                self.wait.until(EC.presence_of_element_located((By.ID, 'UnitId')))
                opts = self.driver.execute_script(
                    """
                    var s = document.getElementById('UnitId');
                    if(!s) return [];
                    return Array.from(s.options).map(o => ({v:o.value, t:(o.text||'').trim()})).filter(x=>x.v);
                    """
                )
                if opts:
                    break
            except Exception:
                opts = []
            time.sleep(0.1)
        units: List[Tuple[str, str]] = []
        for it in opts or []:
            v = it.get('v') or ''
            t = it.get('t') or v
            if v:
                units.append((t, v))
        if not units:
            raise RuntimeError("Не удалось получить список отделов из #UnitId")
        print(f"[UNITS] Найдено отделов: {len(units)}")
        return units

    def choose_unit(self, value: str):
        js = """
        (function(val){
          var s = document.getElementById('UnitId');
          if(!s) return false;
          if (s.value !== val) { s.value = val; }
          try {
            var e=new Event('change',{bubbles:true}); s.dispatchEvent(e);
          } catch(err) {
            var ev=document.createEvent('HTMLEvents'); ev.initEvent('change',true,false); s.dispatchEvent(ev);
          }
          if (window.$ && window.$(s).selectpicker) { try { window.$(s).selectpicker('render'); } catch(e){} }
          return true;
        })(arguments[0]);
        """
        try:
            self.driver.execute_script(js, value)
        except Exception:
            pass
        if self.slow:
            time.sleep(self.slow)

    def set_date_single(self, dt: datetime.date):
        date_str = dt.strftime("%d.%m.%Y")
        js = """
        (function(val){
          function setValue(id){
            var el = document.getElementById(id);
            if(!el) return false;
            el.value = val;
            try { el.dispatchEvent(new Event('input',{bubbles:true})); el.dispatchEvent(new Event('change',{bubbles:true})); } catch(e) {
              var ev=document.createEvent('HTMLEvents'); ev.initEvent('change',true,false); el.dispatchEvent(ev);
            }
            return true;
          }
          var ids = ['Date','DatePeriod','DatePeriodStart','DatePeriodEnd','StartDate','EndDate','SelectedDate'];
          var ok = false;
          for (var i=0;i<ids.length;i++) { ok = setValue(ids[i]) || ok; }
          return ok;
        })(arguments[0]);
        """
        try:
            self.driver.execute_script(js, date_str)
        except Exception:
            pass
        if self.slow:
            time.sleep(self.slow)

    def get_revision_selects(self) -> List[Dict[str, Any]]:
        try:
            selects = self.driver.execute_script(
                """
                function readSelect(s){
                  var opts = Array.from(s.options).map(function(o,idx){
                    return { value:o.value, text:(o.text||'').trim(), index: idx };
                  }).filter(o=>o.value!==undefined && o.value!==null && String(o.value).length>0);
                  return { id: s.id || '', name: s.name || '', options: opts };
                }
                return Array.from(document.querySelectorAll('select'))
                  .filter(s => (s.id||s.name||'').toLowerCase().includes('revision'))
                  .map(readSelect);
                """
            ) or []
        except Exception:
            selects = []
        return selects

    def set_revision_values(self, mapping: Dict[str, str]):
        try:
            self.driver.execute_script(
                """
                (function(map){
                  Object.keys(map||{}).forEach(function(k){
                    var el = document.getElementById(k) || document.querySelector('[name="'+k+'"]');
                    if(!el) return;
                    el.value = map[k];
                    try{ el.dispatchEvent(new Event('change',{bubbles:true})); }catch(e){ var ev=document.createEvent('HTMLEvents'); ev.initEvent('change',true,false); el.dispatchEvent(ev); }
                  });
                })(arguments[0]);
                """,
                mapping,
            )
        except Exception:
            pass

    def build_report(self) -> bool:
        try:
            btn = self.wait.until(EC.element_to_be_clickable((By.ID, 'buildReportButton')))
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            except Exception:
                pass
            btn.click()
            return True
        except Exception:
            try:
                self.driver.execute_script("if (typeof buildReport === 'function') buildReport();")
                return True
            except Exception:
                return False

    def wait_stats_changed(self, old_sig: Optional[int], timeout: int = 30):
        if old_sig is None:
            time.sleep(0.5)
            return
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                new_sig = self.driver.execute_script(
                    "var t=document.querySelector('table'); return t ? (t.innerText||'').length : null;"
                )
                if new_sig and new_sig != old_sig:
                    break
            except Exception:
                pass
            time.sleep(0.05)

    def read_statistics_table(self) -> List[List[str]]:
        rows: List[List[str]] = []
        try:
            table = self.driver.find_element(By.XPATH, "//h3[contains(.,'Статистика')]/following::table[1]")
        except Exception:
            try:
                table = self.driver.find_element(By.CSS_SELECTOR, "table")
            except Exception:
                return rows
        try:
            trs = table.find_elements(By.TAG_NAME, "tr")
        except Exception:
            return rows
        for tr in trs:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue
            rows.append([(td.text or "").strip() for td in tds])
        return rows

    # ---------- CSV ----------
    def reset_csv(self):
        try:
            d = os.path.dirname(self.csv_file)
            if d:
                os.makedirs(d, exist_ok=True)
        except Exception:
            pass
        with open(self.csv_file, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=';')
            w.writerow(["Город","Отдел","Дата","Ревизии","Данные"])

    def collect_table_rows(self, city: str, unit: str, date_str: str, revisions_human: str, rows: List[List[str]]):
        csv_rows: List[List[str]] = []
        sup_rows: List[List[str]] = []
        for r in rows:
            parts = [p.strip() for p in (r or [])]
            metric = parts[0] if parts else ""
            percent = parts[1] if len(parts) > 1 else ""
            amount = parts[2] if len(parts) > 2 else ""
            csv_rows.append([city, unit, date_str, revisions_human, metric, percent, amount])
            sup_rows.append([city, unit, date_str, revisions_human, metric, percent, amount])
        return csv_rows, sup_rows

    def _write_city_rows(self, rows: List[List[str]]):
        if not rows:
            return
        with open(self.csv_file, "a", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=";").writerows(rows)

    def _build_supabase_record(self, row: List[str]) -> Optional[dict]:
        if len(row) < 7:
            return None
        city, dept, date_str, revisions, metric, percent, amount = row[:7]
        fm = self.supabase_field_map
        return {
            fm["city"]: city,
            fm["department"]: dept,
            fm["date"]: self._safe_parse_date(date_str),
            fm["revisions"]: revisions,
            fm["metric"]: metric,
            fm["percent"]: self._parse_number(percent),
            fm["amount"]: self._parse_number(amount),
        }

    def _queue_supabase_rows(self, rows: List[List[str]]) -> None:
        if not rows:
            return
        self._supabase_buffer.extend(rows)
        if len(self._supabase_buffer) >= self.supabase_batch_size:
            self._flush_supabase_buffer()

    def _flush_supabase_buffer(self) -> None:
        if not self._supabase_buffer:
            return
        try:
            self._push_supabase_rows(self._supabase_buffer)
        except Exception as exc:
            print(f"[SUPABASE] Ошибка отправки: {exc}")
        finally:
            self._supabase_buffer.clear()

    def _log_db_response(self, context: str, count: int, stdout: str, stderr: str) -> None:
        try:
            BUGREPORT_FILE.parent.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        try:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with BUGREPORT_FILE.open("a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {context} — {count} rows\n")
                if stdout:
                    f.write(stdout.strip() + "\n")
                if stderr:
                    f.write(f"stderr: {stderr.strip()}\n")
                f.write("\n")
        except Exception as exc:
            print(f"[BUGLOG] Не удалось записать ответ БД: {exc}")

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

    def _parse_number(self, value: str) -> Optional[float]:
        txt = (value or "").strip().replace("\xa0", "").replace(" ", "")
        if not txt:
            return None
        txt = txt.replace(",", ".")
        try:
            return float(txt)
        except ValueError:
            return None

    def _safe_parse_date(self, value: str) -> str:
        txt = (value or "").strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(txt, fmt).date().isoformat()
            except Exception:
                continue
        return txt

    def _post_supabase_chunk(self, chunk: List[dict], allow_conflict: bool = True) -> Dict[str, Any]:
        if not chunk:
            return {"count": 0, "stdout": "", "stderr": "", "endpoint": "", "http_status": 0, "on_conflict": False}
        use_conflict = (
            allow_conflict
            and bool(self.supabase_on_conflict)
            and not self._supabase_conflict_forced_off
        )
        base_url = self.supabase_url.rstrip("/")
        endpoint = f"{base_url}/rest/v1/{self.supabase_table}"
        params = {}
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
        cmd.extend([
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
        ])
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
                retry_info["stdout"] = f"{prev_stdout}\n{extra_msg}".strip() if prev_stdout else extra_msg
                return retry_info
            status_text = f"{http_status}"
            raise RuntimeError(f"Supabase HTTP {status_text}: {error_payload or 'unknown error'}")
        print(f"[SUPABASE] Записано {len(chunk)} записей.")
        return {
            "count": len(chunk),
            "stdout": body_txt,
            "stderr": stderr_txt,
            "endpoint": endpoint,
            "http_status": http_status,
            "on_conflict": use_conflict,
        }

    # ---------- Основной сценарий ----------
    def _parse_date(self, s: str) -> Optional[datetime.date]:
        s = (s or '').strip()
        today = datetime.date.today()
        if not s:
            return today - datetime.timedelta(days=1)
        if s.isdigit():
            if len(s) <= 2:
                try:
                    return datetime.date(today.year, today.month, int(s))
                except Exception:
                    return None
            if len(s) == 8:
                try:
                    return datetime.datetime.strptime(s, "%d%m%Y").date()
                except Exception:
                    return None
        for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except Exception:
                continue
        return None

    def prompt_date_range(self) -> List[datetime.date]:
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
            return [start + datetime.timedelta(days=i) for i in range(days)]

    def run(self):
        self.launch_chrome()
        self.connect_driver()
        dates = self.prompt_date_range()
        if not dates:
            print("[DATES] Диапазон дат пуст — выход.")
            return
        self.reset_csv()

        cities = self.get_cities()
        print(f"[CITIES] К обработке: {[c[0] for c in cities]}")

        for cidx, (city_name, city_uuid) in enumerate(cities, start=1):
            print("\n" + "#" * 80)
            print(f"[CITY] ({cidx}/{len(cities)}) {city_name}")
            try:
                self.select_city(city_uuid)
                self.open_report_for_city(city_uuid)

                units = self.get_units()
                for uidx, (unit_name, unit_val) in enumerate(units, start=1):
                    print("\n" + "=" * 80)
                    print(f"[UNIT] ({uidx}/{len(units)}) {unit_name}")
                    self.choose_unit(unit_val)
                    unit_csv_rows: List[List[str]] = []
                    unit_sup_rows: List[List[str]] = []
                    rev_selects = self.get_revision_selects()
                    options_lists = []
                    for s in rev_selects:
                        opts = s.get('options') or []
                        if len(opts) <= 1:
                            options_lists.append([{ 'select': s.get('id') or s.get('name'), 'value': (opts[0]['value'] if opts else ''), 'text': (opts[0]['text'] if opts else '') }])
                        else:
                            options_lists.append([{ 'select': s.get('id') or s.get('name'), 'value': o['value'], 'text': o['text'] } for o in opts])
                    if not options_lists:
                        options_lists = [[{ 'select': '', 'value': '', 'text': '' }]]
                    base_options = options_lists

                    for dt_ in dates:
                        date_str = dt_.strftime("%d.%m.%Y")
                        self.set_date_single(dt_)
                        options_lists = []
                        for items in base_options:
                            fresh = []
                            for item in items:
                                select_id = item.get('select') or ''
                                value = item.get('value') or ''
                                text = item.get('text') or value
                                fresh.append({'select': select_id, 'value': value, 'text': text})
                            if not fresh:
                                fresh = [{'select': '', 'value': '', 'text': ''}]
                            options_lists.append(fresh)
                        combos = list(itertools.product(*options_lists))
                        if len(combos) > 25:
                            print(f"[REVISIONS] Слишком много комбинаций ({len(combos)}), обрезаю до 25…")
                            combos = combos[:25]

                        for ridx, combo in enumerate(combos, start=1):
                            combo_map = {}
                            revisions_human_list: List[str] = []
                            for item in combo:
                                sid = item.get('select') or ''
                                val = item.get('value') or ''
                                txt = item.get('text') or val
                                if sid:
                                    combo_map[sid] = val
                                    revisions_human_list.append(f"{sid}:{txt}")
                            revisions_human = ", ".join([x for x in revisions_human_list if x]) or "(нет)"
                            if combo_map:
                                self.set_revision_values(combo_map)
                            try:
                                old_sig = self.driver.execute_script(
                                    "var t=document.querySelector('table'); return t ? (t.innerText||'').length : null;"
                                )
                            except Exception:
                                old_sig = None
                            ok = self.build_report()
                            if not ok:
                                print("[BUILD] Не удалось нажать кнопку построения.")
                            self.wait_stats_changed(old_sig, timeout=30)

                            rows = self.read_statistics_table()
                            print(f"[CSV] {city_name} / {unit_name} / {date_str} / rev#{ridx}: {len(rows)} строк")
                            if rows:
                                csv_rows, sup_rows = self.collect_table_rows(city_name, unit_name, date_str, revisions_human, rows)
                                unit_csv_rows.extend(csv_rows)
                                unit_sup_rows.extend(sup_rows)

                    if unit_csv_rows:
                        self._write_city_rows(unit_csv_rows)
                    if unit_sup_rows:
                        try:
                            responses = self._push_supabase_rows(unit_sup_rows)
                        except Exception as exc:
                            total_rows = len(unit_sup_rows)
                            error_text = f"Ошибка отправки в Supabase: {exc}"
                            print(f"[SUPABASE] {error_text}")
                            self._log_db_response(f"{city_name} / {unit_name} — ERROR (total rows: {total_rows})", 0, "", error_text)
                            responses = []
                        if responses:
                            total_rows = len(unit_sup_rows)
                            for resp_idx, resp in enumerate(responses, start=1):
                                context = f"{city_name} / {unit_name} — chunk {resp_idx}/{len(responses)} (total rows: {total_rows})"
                                self._log_db_response(context, int(resp.get("count", 0)), str(resp.get("stdout", "")), str(resp.get("stderr", "")))
                    self.refresh_report_page()

                self._flush_supabase_buffer()

            except Exception as e:
                print(f"[WARN] Ошибка в городе {city_name}: {e}")
            finally:
                try:
                    self.back_to_select_role()
                except Exception:
                    pass

        self._flush_supabase_buffer()

        print(f"[DONE] Готово! Файл {self.csv_file} сохранён.")

    def close(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass

    @staticmethod
    def _wait_port(port: int, timeout: int = 10) -> bool:
        for _ in range(timeout * 10):
            with socket.socket() as s:
                if s.connect_ex(("127.0.0.1", port)) == 0:
                    return True
            time.sleep(0.1)
        return False


if __name__ == "__main__":
    bot = LossesAndExcessReporter()
    try:
        bot.run()
    finally:
        bot.close()
