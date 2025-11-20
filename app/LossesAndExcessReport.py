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
import json
import re
import urllib.parse


PORT = 9223
CSV_FILE = "reports/LossesAndExcessReport.csv"
JSON_FILE = "reports/LossesAndExcessReport.jsonl"
BUGREPORT_FILE = Path("reports/bugrepot.txt")
REPORT_URL = "https://officemanager.dodopizza.ru/InventoryManager/LossesAndExcees"
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ROLE_ID = os.environ.get("ROLE_ID", "7")
SLOW_DELAY = float(os.environ.get("SLOW_DELAY", "0"))
STEALTH = os.environ.get("STEALTH", "1")
SUPABASE_CONFIG_FILE = Path(os.environ.get("SUPABASE_CONFIG_FILE", "config/api"))
SUPABASE_KEY_FILE = Path(os.environ.get("SUPABASE_KEY_FILE", "config/api.key"))

LIQUID_KEYWORDS = tuple(k.lower() for k in ("сок", "вода", "смузи", "чай", "морс", "лимонад", "компот"))
LIQUID_KEYWORD_EXCEPTIONS = tuple(k.lower() for k in ("бонаква", "бон аква", "bone aqua", "bonaqua"))
TARGET_METRIC_KEY = "сыр моцарелла"
ESSENTUKI_KEYWORDS = ("ессентук",)
DEFAULT_HIGHLIGHT_BG = "#d4f8c4"
DEFAULT_HIGHLIGHT_BORDER = "2px solid #38a169"


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
        self.supabase_queue_table = os.environ.get("SUPABASE_QUEUE_TABLE", "metric_queue").strip() or "metric_queue"
        self.supabase_queue_filter = os.environ.get("SUPABASE_QUEUE_FILTER", "is_done=eq.false&order=created_at.asc").strip()
        self._supabase_warned = False
        self._supabase_conflict_forced_off = False
        self._supabase_conflict_notice_shown = False
        self.highlight_bg = os.environ.get("HIGHLIGHT_BG", DEFAULT_HIGHLIGHT_BG)
        self.highlight_border = os.environ.get("HIGHLIGHT_BORDER", DEFAULT_HIGHLIGHT_BORDER)
        self.highlight_target_bg = os.environ.get("HIGHLIGHT_TARGET_BG", "#c6f6d5")
        self.highlight_target_border = os.environ.get("HIGHLIGHT_TARGET_BORDER", "2px solid #38a169")
        self.highlight_liquid_bg = os.environ.get("HIGHLIGHT_LIQUID_BG", "#fde8e8")
        self.highlight_liquid_border = os.environ.get("HIGHLIGHT_LIQUID_BORDER", "2px solid #f56565")
        self._city_cache: Dict[str, Tuple[str, str]] = {}
        default_field_map = {
            "city": os.environ.get("SUPABASE_FIELD_CITY", "city"),
            "department": os.environ.get("SUPABASE_FIELD_DEPARTMENT", "department"),
            "date": os.environ.get("SUPABASE_FIELD_DATE", "dt"),
            "revisions": os.environ.get("SUPABASE_FIELD_REVISIONS", "revisions"),
            "metric": os.environ.get("SUPABASE_FIELD_METRIC", "metric_name"),
            "percent": os.environ.get("SUPABASE_FIELD_PERCENT", "percent_value"),
            "amount": os.environ.get("SUPABASE_FIELD_AMOUNT", "amount_value"),
            "queue_id": os.environ.get("SUPABASE_FIELD_QUEUE_ID", "id_queue"),
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

    def _highlight_elements(self, selectors: List[str]):
        selector = ",".join([s for s in selectors if s]).strip()
        if not selector:
            return
        try:
            self.driver.execute_script(
                """
                (function(sel, bg, border){
                  if(!sel) return;
                  var nodes;
                  try { nodes = document.querySelectorAll(sel); } catch(err){ return; }
                  nodes.forEach(function(el){
                    try {
                      if(bg){ el.style.setProperty('background-color', bg, 'important'); }
                      if(border){ el.style.setProperty('border', border, 'important'); }
                    } catch(styleErr){}
                  });
                })(arguments[0], arguments[1], arguments[2]);
                """,
                selector,
                self.highlight_bg,
                self.highlight_border,
            )
        except Exception:
            pass

    def _highlight_webelement(self, element, bg: Optional[str] = None, border: Optional[str] = None):
        if not element:
            return
        try:
            self.driver.execute_script(
                """
                (function(el, bg, border){
                  if(!el) return;
                  try {
                    if(bg){ el.style.setProperty('background-color', bg, 'important'); }
                    if(border){ el.style.setProperty('border', border, 'important'); }
                  } catch(err){}
                })(arguments[0], arguments[1], arguments[2]);
                """,
                element,
                bg,
                border,
            )
        except Exception:
            pass

    def _normalize_city_key(self, name: str) -> str:
        base = re.sub(r"[\s\-]+", " ", (name or "").strip()).lower()
        base = re.sub(r"\s+\d+$", "", base).strip()
        return base

    def _normalize_unit_key(self, name: str) -> str:
        return re.sub(r"[\s\-]+", "", (name or "").strip().lower())

    def _split_pizzeria_name(self, name: str) -> Tuple[str, str]:
        name = (name or "").strip()
        match = re.match(r"^(.*?)(\s+(\d+))?$", name)
        if not match:
            return name, name
        base = (match.group(1) or "").strip()
        number = (match.group(3) or "").strip()
        if number:
            unit = f"{base}-{number}"
        else:
            unit = base
        return base, unit

    def _ensure_city_cache(self):
        if self._city_cache:
            return
        cities = self.get_cities()
        for name, uuid in cities:
            key = self._normalize_city_key(name)
            if key and key not in self._city_cache:
                self._city_cache[key] = (name, uuid)

    def _find_city_entry(self, target_name: str) -> Optional[Tuple[str, str]]:
        self._ensure_city_cache()
        key = self._normalize_city_key(target_name)
        return self._city_cache.get(key)

    def _find_unit_option(self, units: List[Tuple[str, str]], target_name: str) -> Optional[Tuple[str, str]]:
        target_key = self._normalize_unit_key(target_name)
        alt_key = self._normalize_unit_key(target_name.replace("-", " "))
        for display, value in units:
            unit_key = self._normalize_unit_key(display)
            if unit_key == target_key or unit_key == alt_key:
                return display, value
        return None

    def _fetch_metric_queue(self) -> List[Dict[str, Any]]:
        if not self.supabase_enabled:
            print("[QUEUE] Supabase не настроен — задачи не получены.")
            return []
        base_url = self.supabase_url.rstrip("/")
        query_parts = ["select=*"]
        if self.supabase_queue_filter:
            query_parts.append(self.supabase_queue_filter)
        endpoint = f"{base_url}/rest/v1/{self.supabase_queue_table}?{'&'.join(query_parts)}"
        cmd = [
            "curl",
            "-sS",
            "-H",
            f"apikey: {self.supabase_key}",
            "-H",
            f"Authorization: Bearer {self.supabase_key}",
            endpoint,
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        except Exception as exc:
            print(f"[QUEUE] Не удалось выполнить curl: {exc}")
            return []
        if result.returncode != 0:
            print(f"[QUEUE] curl exit code {result.returncode}: {result.stderr.strip()}")
            return []
        try:
            payload = json.loads(result.stdout or "[]")
        except json.JSONDecodeError as exc:
            print(f"[QUEUE] Некорректный JSON: {exc}")
            return []
        if not isinstance(payload, list):
            print("[QUEUE] Ожидался список задач.")
            return []
        return payload

    def _match_revision_values(
        self,
        selects: List[Dict[str, Any]],
        revisions: List[str],
    ) -> Tuple[Dict[str, str], List[str]]:
        assignments: Dict[str, str] = {}
        missing: List[str] = []
        used_selects: set[str] = set()

        for revision_text in revisions:
            revision_text = (revision_text or "").strip()
            if not revision_text:
                continue
            revision_key = revision_text.lower()
            matched = False
            for select_info in selects:
                select_id = (select_info.get("id") or "").strip()
                select_name = (select_info.get("name") or "").strip()
                identifier = select_id or select_name
                if not identifier or identifier in used_selects:
                    continue
                for opt in select_info.get("options") or []:
                    opt_text = (opt.get("text") or "").strip()
                    if revision_key in opt_text.lower():
                        assignments[identifier] = opt.get("value") or ""
                        used_selects.add(identifier)
                        matched = True
                        break
                if matched:
                    break
            if not matched:
                missing.append(revision_text)
        return assignments, missing

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
        self._highlight_elements(["#UnitId", "[name='UnitId']", "select[data-live-search][name='UnitId']"])
        if self.slow:
            time.sleep(self.slow)

    def set_date_single(self, dt: datetime.date):
        self.set_date_range(dt, dt)

    def _read_date_values(self) -> Dict[str, str]:
        try:
            values = self.driver.execute_script(
                """
                var out = {};
                var nodes = document.querySelectorAll('[id*=\"Date\"],[name*=\"Date\"],[id*=\"Period\"],[name*=\"Period\"]');
                Array.prototype.forEach.call(nodes, function(el){
                  if(!el) return;
                  var key = el.id || el.name;
                  if(!key) return;
                  if(typeof el.value !== 'undefined'){ out[key] = String(el.value || ''); }
                });
                return out;
                """
            )
            if isinstance(values, dict):
                return {str(k): str(v) for k, v in values.items()}
        except Exception:
            pass
        return {}

    def set_date_range(self, start: datetime.date, end: datetime.date):
        start_str = start.strftime("%d.%m.%Y")
        end_str = end.strftime("%d.%m.%Y")
        expected_range = f"{start_str} - {end_str}"
        js = """
        (function(startVal, endVal, highlightColor, highlightBorder){
          function toDate(val){
            var parts = (val||'').split('.');
            if(parts.length === 3){
              return new Date(parts[2], parseInt(parts[1],10)-1, parts[0]);
            }
            return null;
          }
          function setElementValue(el, val){
            if(!el) return;
            try {
              if (typeof el.value !== 'undefined' && el.value !== val) {
                el.value = val;
              }
              if (el.setAttribute) {
                el.setAttribute('value', val);
                if (el.getAttribute('data-value') !== null) {
                  el.setAttribute('data-value', val);
                }
              }
              if (el.type === 'date' && el.valueAsDate !== undefined) {
                var d = toDate(val);
                if (d) { el.valueAsDate = d; }
              }
              if (window.$) {
                try {
                  window.$(el).val(val);
                  if (window.$(el).datepicker) { window.$(el).datepicker('setDate', val); }
                  window.$(el).trigger('change');
                } catch(jqErr){}
              }
              try { el.dispatchEvent(new Event('input',{bubbles:true})); } catch(evtErr1){
                try { var ev1=document.createEvent('HTMLEvents'); ev1.initEvent('input',true,false); el.dispatchEvent(ev1); } catch(_) {}
              }
              try { el.dispatchEvent(new Event('change',{bubbles:true})); } catch(evtErr2){
                try { var ev2=document.createEvent('HTMLEvents'); ev2.initEvent('change',true,false); el.dispatchEvent(ev2); } catch(_) {}
              }
              try { el.dispatchEvent(new Event('blur',{bubbles:true})); } catch(_) {}
              try {
                if (el.style && highlightColor) { el.style.setProperty('background-color', highlightColor, 'important'); }
                if (el.style && highlightBorder) { el.style.setProperty('border', highlightBorder, 'important'); }
              } catch(styleErr){}
            } catch(err){}
          }
          function setSelectors(selectors, val){
            var applied = false;
            (selectors||[]).forEach(function(sel){
              try {
                var nodes = document.querySelectorAll(sel);
                Array.prototype.forEach.call(nodes, function(node){
                  setElementValue(node, val);
                  applied = true;
                });
              } catch(err){}
            });
            return applied;
          }
          var startSelectors = [
            '#DatePeriodStart','#StartDate','#DateStart','#DefaultBeginDateString',
            'input[name=\"DatePeriodStart\"]','input[name=\"StartDate\"]','input[name=\"DateStart\"]',
            'input[name=\"DefaultBeginDateString\"]','input[id$=\"BeginDateString\"]','input[data-role=\"start-date\"]'
          ];
          var endSelectors = [
            '#DatePeriodEnd','#EndDate','#DateEnd','#Date','#SelectedDate','#DefaultEndDateString',
            'input[name=\"DatePeriodEnd\"]','input[name=\"EndDate\"]','input[name=\"DateEnd\"]','input[name=\"Date\"]','input[name=\"SelectedDate\"]',
            'input[name=\"DefaultEndDateString\"]','input[id$=\"EndDateString\"]','input[data-role=\"end-date\"]'
          ];
          var rangeSelectors = [
            '#DatePeriod','#DateRange','#Period','#DefaultPeriodString',
            'input[name=\"DatePeriod\"]','input[name=\"DateRange\"]','input[name=\"Period\"]','input[name=\"DefaultPeriodString\"]'
          ];
          var singleSelectors = [
            '#Date','#SelectedDate','input[name=\"Date\"]','input[name=\"SelectedDate\"]','input[id$=\"SelectedDateString\"]'
          ];
          var rangeVal = startVal + ' - ' + endVal;
          var anyApplied = false;
          anyApplied = setSelectors(startSelectors, startVal) || anyApplied;
          anyApplied = setSelectors(endSelectors, endVal) || anyApplied;
          anyApplied = setSelectors(rangeSelectors, rangeVal) || anyApplied;
          anyApplied = setSelectors(singleSelectors, endVal) || anyApplied;
          return anyApplied;
        })(arguments[0], arguments[1], arguments[2], arguments[3]);
        """
        applied = False
        for attempt in range(3):
            try:
                self.driver.execute_script(js, start_str, end_str, self.highlight_bg, self.highlight_border)
            except Exception:
                pass
            if self.slow:
                time.sleep(self.slow)
            values = self._read_date_values()
            start_val = (
                values.get("DatePeriodStart")
                or values.get("StartDate")
                or values.get("DateStart")
                or values.get("DefaultBeginDateString")
                or next((v for k, v in values.items() if k.lower().endswith("begindatestring")), "")
            )
            end_val = (
                values.get("DatePeriodEnd")
                or values.get("EndDate")
                or values.get("DateEnd")
                or values.get("Date")
                or values.get("SelectedDate")
                or values.get("DefaultEndDateString")
                or next((v for k, v in values.items() if k.lower().endswith("enddatestring")), "")
            )
            range_val = (
                values.get("DatePeriod")
                or values.get("DateRange")
                or values.get("Period")
                or values.get("DefaultPeriodString")
                or next((v for k, v in values.items() if k.lower().endswith("periodstring")), "")
            )
            if (start_val == start_str and end_val == end_str) or (range_val == expected_range):
                applied = True
                break
            time.sleep(0.1)
        if not applied:
            fallback = range_val or ((start_val or "?") + "…" + (end_val or "?"))
            print(f"[DATES] Не удалось надёжно установить период {expected_range}, текущее значение: {fallback}")
        else:
            print(f"[DATES] Период установлен: {expected_range}")

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
        highlight_selectors: List[str] = []
        for sel in selects or []:
            sid = (sel.get("id") or "").strip()
            sname = (sel.get("name") or "").strip()
            if sid:
                safe_id = sid.replace("'", "\\'")
                highlight_selectors.append(f"[id='{safe_id}']")
            if sname:
                safe_name = sname.replace("'", "\\'")
                highlight_selectors.append(f"[name='{safe_name}']")
        if highlight_selectors:
            self._highlight_elements(highlight_selectors)
        return selects

    def set_revision_values(self, mapping: Dict[str, str]):
        try:
            self.driver.execute_script(
                """
                (function(map, highlightColor, highlightBorder){
                  Object.keys(map||{}).forEach(function(k){
                    var el = document.getElementById(k) || document.querySelector('[name="'+k+'"]');
                    if(!el) return;
                    el.value = map[k];
                    try{ el.dispatchEvent(new Event('change',{bubbles:true})); }catch(e){ var ev=document.createEvent('HTMLEvents'); ev.initEvent('change',true,false); el.dispatchEvent(ev); }
                    try {
                      if (el.style && highlightColor) { el.style.setProperty('background-color', highlightColor, 'important'); }
                      if (el.style && highlightBorder) { el.style.setProperty('border', highlightBorder, 'important'); }
                    } catch(err){}
                  });
                })(arguments[0], arguments[1], arguments[2]);
                """,
                mapping,
                self.highlight_bg,
                self.highlight_border,
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
            self._highlight_webelement(btn, self.highlight_bg, self.highlight_border)
            btn.click()
            self._pause_after_build()
            return True
        except Exception:
            try:
                self._highlight_elements(["#buildReportButton", "[id='buildReportButton']", "button[onclick*='buildReport']", "a[onclick*='buildReport']"])
                self.driver.execute_script("if (typeof buildReport === 'function') buildReport();")
                self._pause_after_build()
                return True
            except Exception:
                return False

    def _pause_after_build(self):
        pass

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
        self._highlight_webelement(table, self.highlight_bg, self.highlight_border)
        target_kw = TARGET_METRIC_KEY.lower()
        for tr in trs:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue
            rows.append([(td.text or "").strip() for td in tds])
            try:
                metric_text = (tds[0].text or "").strip()
            except Exception:
                metric_text = ""
            metric_lower = metric_text.lower()
            if self._is_target_metric(metric_lower):
                self._highlight_webelement(tr, self.highlight_target_bg, self.highlight_target_border)
            elif self._is_liquid_metric(metric_lower):
                self._highlight_webelement(tr, self.highlight_liquid_bg, self.highlight_liquid_border)
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

    def reset_json(self):
        try:
            d = os.path.dirname(JSON_FILE)
            if d:
                os.makedirs(d, exist_ok=True)
        except Exception:
            pass
        Path(JSON_FILE).write_text("", encoding="utf-8")

    def reset_outputs(self):
        self.reset_csv()
        self.reset_json()

    def collect_table_rows(
        self,
        city: str,
        unit: str,
        date_str: str,
        revisions_human: str,
        rows: List[List[str]],
        queue_id: Optional[Any] = None,
    ):
        csv_rows: List[List[str]] = []
        sup_rows: List[List[str]] = []
        for r in rows:
            parts = [p.strip() for p in (r or [])]
            metric = parts[0] if parts else ""
            percent = parts[1] if len(parts) > 1 else ""
            amount = parts[2] if len(parts) > 2 else ""
            csv_rows.append([city, unit, date_str, revisions_human, metric, percent, amount])
            sup_row = [city, unit, date_str, revisions_human, metric, percent, amount]
            if queue_id is not None:
                sup_row.append(queue_id)
            sup_rows.append(sup_row)
        return csv_rows, sup_rows

    def _write_city_rows(self, rows: List[List[str]]):
        if not rows:
            return
        with open(self.csv_file, "a", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=";").writerows(rows)

    def _write_json_row(self, csv_row: List[str], queue_id: Optional[Any]) -> None:
        if not csv_row:
            return
        try:
            record = {
                "city": csv_row[0] if len(csv_row) > 0 else "",
                "department": csv_row[1] if len(csv_row) > 1 else "",
                "date": csv_row[2] if len(csv_row) > 2 else "",
                "revisions": csv_row[3] if len(csv_row) > 3 else "",
                "metric": csv_row[4] if len(csv_row) > 4 else "",
                "percent": csv_row[5] if len(csv_row) > 5 else "",
                "amount": csv_row[6] if len(csv_row) > 6 else "",
                "queue_id": queue_id,
            }
            with open(JSON_FILE, "a", encoding="utf-8") as f:
                f.write(json.dumps(record, ensure_ascii=False))
                f.write("\n")
        except Exception as exc:
            print(f"[JSON] Не удалось записать запись: {exc}")

    def _build_supabase_record(self, row: List[str]) -> Optional[dict]:
        if len(row) < 7:
            return None
        city, dept, date_str, revisions, metric, percent, amount = row[:7]
        queue_id = row[7] if len(row) > 7 else None
        fm = self.supabase_field_map
        return {
            fm["city"]: city,
            fm["department"]: dept,
            fm["date"]: self._safe_parse_date(date_str),
            fm["revisions"]: revisions,
            fm["metric"]: metric,
            fm["percent"]: self._parse_number(percent),
            fm["amount"]: self._parse_number(amount),
            fm.get("queue_id", "id_queue"): queue_id,
        }

    def _queue_supabase_rows(self, rows: List[List[str]]) -> None:
        if not rows:
            return
        try:
            self._push_supabase_rows(rows)
        except Exception as exc:
            print(f"[SUPABASE] Ошибка отправки: {exc}")

    def _flush_supabase_buffer(self) -> None:
        return

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

    def _mark_queue_done(self, queue_id: Optional[Any]) -> None:
        if not self.supabase_enabled:
            return
        if queue_id is None:
            return
        queue_id_str = str(queue_id).strip()
        if not queue_id_str:
            return
        base_url = self.supabase_url.rstrip("/")
        encoded_id = urllib.parse.quote(queue_id_str, safe="")
        endpoint = f"{base_url}/rest/v1/{self.supabase_queue_table}?id=eq.{encoded_id}"
        payload = json.dumps({"is_done": True})
        cmd = [
            "curl",
            "-sS",
            "-X",
            "PATCH",
            "-H",
            f"apikey: {self.supabase_key}",
            "-H",
            f"Authorization: Bearer {self.supabase_key}",
            "-H",
            "Content-Type: application/json",
            "-H",
            "Prefer: return=minimal",
            endpoint,
            "-d",
            payload,
        ]
        if self.supabase_timeout > 0:
            cmd.insert(1, "--max-time")
            cmd.insert(2, str(self.supabase_timeout))
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "").strip()
            print(f"[QUEUE] Не удалось обновить is_done для id={queue_id_str}: {err}")

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

    def _is_liquid_metric(self, metric_lower: str) -> bool:
        if not metric_lower:
            return False
        for exception in LIQUID_KEYWORD_EXCEPTIONS:
            if exception and exception in metric_lower:
                return False
        for keyword in LIQUID_KEYWORDS:
            if keyword and keyword in metric_lower:
                return True
        return False

    def _is_target_metric(self, metric_lower: str) -> bool:
        return bool(metric_lower and TARGET_METRIC_KEY.lower() in metric_lower)

    def _process_task(self, task: Dict[str, Any]):
        pizzeria_name = task.get("pizzeria_name") or ""
        date_start_raw = task.get("date_start") or ""
        date_end_raw = task.get("date_end") or ""
        revision_start = task.get("revision_start_string") or ""
        revision_end = task.get("revision_end_string") or ""
        queue_id = task.get("id")

        city_name, unit_candidate = self._split_pizzeria_name(pizzeria_name)
        city_entry = self._find_city_entry(city_name)
        if not city_entry:
            raise RuntimeError(f"Город '{city_name}' не найден на странице выбора.")
        city_display, city_uuid = city_entry

        try:
            start_dt = datetime.date.fromisoformat(date_start_raw)
        except Exception:
            start_dt = self._parse_date(date_start_raw) or datetime.date.today()
        try:
            end_dt = datetime.date.fromisoformat(date_end_raw)
        except Exception:
            end_dt = self._parse_date(date_end_raw) or datetime.date.today()

        self.select_city(city_uuid)
        self.open_report_for_city(city_uuid)

        units = self.get_units()
        unit_match = self._find_unit_option(units, unit_candidate)
        if not unit_match:
            raise RuntimeError(f"Отдел '{unit_candidate}' не найден в списке для города {city_display}.")
        unit_display, unit_value = unit_match
        self.choose_unit(unit_value)

        self.set_date_range(start_dt, end_dt)
        rev_selects = self.get_revision_selects()
        assignments, missing = self._match_revision_values(rev_selects, [revision_start, revision_end])
        if assignments:
            self.set_revision_values(assignments)
        if missing:
            print(f"[REVISIONS] Не найдены ревизии: {', '.join(missing)}")

        try:
            old_sig = self.driver.execute_script(
                "var t=document.querySelector('table'); return t ? (t.innerText||'').length : null;"
            )
        except Exception:
            old_sig = None

        if not self.build_report():
            print("[BUILD] Не удалось инициировать построение отчёта.")
        self.wait_stats_changed(old_sig, timeout=30)

        rows = self.read_statistics_table()
        date_label = end_dt.strftime("%d.%m.%Y")
        revision_label = f"{revision_start} -> {revision_end}".strip(" ->")
        print(f"[DATA] {city_display} / {unit_display} / {date_label}: {len(rows)} строк")
        if not rows:
            self._mark_queue_done(queue_id)
            return

        csv_rows, sup_rows = self.collect_table_rows(
            city_display,
            unit_display,
            date_label,
            revision_label,
            rows,
            queue_id=queue_id,
        )

        for csv_row, sup_row in zip(csv_rows, sup_rows):
            self._write_city_rows([csv_row])
            queue_value = sup_row[7] if len(sup_row) > 7 else queue_id
            self._write_json_row(csv_row, queue_value)
            self._queue_supabase_rows([sup_row])

        self._mark_queue_done(queue_id)

    def run(self):
        self.launch_chrome()
        self.connect_driver()
        tasks = self._fetch_metric_queue()
        if not tasks:
            print("[QUEUE] Задач не найдено — выхожу.")
            return
        print(f"[QUEUE] Получено задач: {len(tasks)}")
        self.reset_outputs()
        self._ensure_city_cache()

        for idx, task in enumerate(tasks, start=1):
            print("\n" + "#" * 80)
            print(f"[TASK] ({idx}/{len(tasks)}) {task.get('pizzeria_name', 'unknown')} — {task.get('date_start')} -> {task.get('date_end')}")
            try:
                self._process_task(task)
            except Exception as exc:
                print(f"[WARN] Ошибка обработки задачи #{task.get('id')}: {exc}")
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
