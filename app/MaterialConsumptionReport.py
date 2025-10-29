from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from typing import Optional, List, Tuple, Dict, Any

import csv
import datetime
import json
import os
from pathlib import Path
import socket
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request


# Конфигурация по умолчанию (страницы и роль зашиты в коде)
PORT = 9222
CSV_FILE = "reports/MaterialConsumptionReport.csv"
REPORT_URL = "https://officemanager.dodopizza.ru/OfficeManager/MaterialConsumption"
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ROLE_ID = "7"  # роль Офис‑менеджера
SLOW_DELAY = float(os.environ.get("SLOW_DELAY", "0"))
STEALTH = os.environ.get("STEALTH", "1")  # 1 — включить анти‑headless твики
SUPABASE_CONFIG_FILE = Path(os.environ.get("SUPABASE_CONFIG_FILE", "config/api"))
BUGREPORT_FILE = Path("reports/bugrepot.txt")
REPO_ROOT = Path(__file__).resolve().parent.parent
LOCAL_CHROMEDRIVER = REPO_ROOT / "bin" / "chromedriver"


def log_bugreport(context: str, count: int, stdout: str, stderr: str) -> None:
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


class OfficeMaterialConsumptionReporter:
    """Сбор данных по отчёту MaterialConsumption."""

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
        self.supabase_table = os.environ.get("SUPABASE_TABLE", "material_consumption").strip() or "material_consumption"
        self.supabase_on_conflict = os.environ.get("SUPABASE_ON_CONFLICT", "").strip()
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
            "category": os.environ.get("SUPABASE_FIELD_CATEGORY", "category"),
            "sales": os.environ.get("SUPABASE_FIELD_SALES", "sales"),
            "production": os.environ.get("SUPABASE_FIELD_PRODUCTION", "production"),
            "staff_food": os.environ.get("SUPABASE_FIELD_STAFF_FOOD", "staff_meals"),
            "cancellation": os.environ.get("SUPABASE_FIELD_CANCELLATION", "cancellations"),
            "defect": os.environ.get("SUPABASE_FIELD_DEFECT", "defects"),
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

    # ---------- Инициализация браузера ----------
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
                print("[INIT] Chrome.exe не найден, пропускаю внешний запуск.")
        else:
            print("[INIT] Linux/Docker: внешний Chrome не запускаю (использую драйвер).")

    @staticmethod
    def _chromedriver_binary_ok(path: str) -> bool:
        if not os.path.exists(path) or not os.access(path, os.X_OK):
            return False
        try:
            with open(path, "rb") as fh:
                signature = fh.read(4)
        except OSError:
            return False
        # ELF — Linux, MZ — Windows PE. Обёртки shell/powershell начинаются с "#!" и не подходят.
        return signature.startswith(b"\x7fELF") or signature.startswith(b"MZ")

    def _make_service(self) -> Service:
        env_path = os.environ.get("CHROMEDRIVER", "").strip()
        candidates = [env_path, str(LOCAL_CHROMEDRIVER), "/usr/bin/chromedriver"]
        for path in candidates:
            if not path:
                continue
            if self._chromedriver_binary_ok(path):
                print(f"[DRIVER] Использую chromedriver: {path}")
                return Service(path)
            if os.path.exists(path):
                print(f"[DRIVER] {path} найден, но не выглядит как исполняемый chromedriver — пропускаю.")
        if env_path:
            print(f"[DRIVER] CHROMEDRIVER={env_path}, но рабочий бинарник не найден. Скачаю свежую копию…")
        return Service(ChromeDriverManager().install())

    def connect_driver(self):
        print("[DRIVER] Инициализация драйвера Chrome…")
        options = webdriver.ChromeOptions()
        # Лёгкие твики ещё на уровне опций
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
            if user_dir:
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

        # Дополнительные анти‑headless настройки через CDP/JS
        if STEALTH == "1":
            self._apply_stealth()

    # ---------- Навигация и авторизация ----------
    def open_select_department(self):
        print("[NAV] Перехожу на экран выбора города…")
        for attempt in range(2):
            try:
                self.driver.get(SELECT_DEPARTMENT_URL)
            except Exception:
                pass
            self.ensure_role_selected()
            if "/SelectDepartment" in (self.driver.current_url or ""):
                return
            if attempt == 0:
                # Попробуем вернуться на SelectRole и сбросить привязку
                try:
                    self.driver.get(BACK_TO_SELECT_ROLE_URL)
                except Exception:
                    pass
                try:
                    WebDriverWait(self.driver, 10).until(EC.url_contains("/SelectRole"))
                except Exception:
                    pass
        print(f"[NAV] Не удалось открыть SelectDepartment, текущий URL: {self.driver.current_url}")

    def choose_role(self):
        print(f"[AUTH] Выбираю роль {ROLE_ID}…")
        # Подсказка: выведем доступные роли, чтобы было проще подобрать ROLE_ID
        try:
            roles = self.driver.execute_script(
                """
                return Array.from(document.querySelectorAll('[name=\"roleId\"]'))
                  .map(el => ({ value: el.getAttribute('value') || '', text: (el.textContent||el.value||'').trim() }));
                """
            ) or []
            if roles:
                print("[AUTH] Доступные роли:")
                for r in roles:
                    print(f"[AUTH] value={r.get('value')} text={r.get('text')}")
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

    def ensure_role_selected(self, city_uuid: Optional[str] = None):
        if "/SelectRole" in self.driver.current_url:
            self.choose_role()
            if city_uuid:
                try:
                    self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="uuid"][value="{city_uuid}"]'))).click()
                except Exception:
                    pass

    # ---------- Города ----------
    def get_cities(self) -> List[Tuple[str, str]]:
        print("[CITIES] Собираю список городов…")
        self.open_select_department()
        # Быстрый лог текущего URL для диагностики
        try:
            print(f"[DEBUG] current_url: {self.driver.current_url}")
        except Exception:
            pass
        try:
            self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'button[name="uuid"], a[name="uuid"]')))
        except Exception:
            self._debug_dump("cities-timeout")
            raise
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
        print("[NAV] Перехожу на страницу MaterialConsumption…")
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

    # ---------- Отделы и фильтры ----------
    def get_departments(self) -> List[str]:
        print("[DEPTS] Получаю список отделов…")
        names: List[str] = []
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
        if not names:
            raise RuntimeError("Список отделов пуст (SelectedUnitIds)")
        return names

    def choose_department(self, name: str):
        js = """
        (function(targetText){
          var s = document.getElementById('SelectedUnitIds');
          if(!s) return false;
          var changed = false;
          for (var i=0; i<s.options.length; i++) {
            var o = s.options[i];
            var sel = ((o.text||'').trim() === targetText);
            if (o.selected !== sel) { o.selected = sel; changed = true; }
          }
          if (changed) {
            var e; try{ e=new Event('change',{bubbles:true}); } catch(err){ e=document.createEvent('HTMLEvents'); e.initEvent('change',true,false); }
            s.dispatchEvent(e);
            if (window.$ && window.$(s).selectpicker) { try { window.$(s).selectpicker('render'); } catch(e){} }
          }
          return true;
        })(arguments[0]);
        """
        try:
            self.driver.execute_script(js, name)
        except Exception:
            pass
        if self.slow:
            time.sleep(self.slow)

    # ---------- Построение и чтение отчёта ----------
    def build_for_date(self, dt: datetime.date):
        date_str = dt.strftime("%d.%m.%Y")
        # Тип представления: период
        try:
            self.driver.execute_script(
                "var s=document.getElementById('CurrentViewType'); if(s){ s.value='Full'; var e; try{e=new Event('change',{bubbles:true});}catch(err){e=document.createEvent('HTMLEvents'); e.initEvent('change',true,false);} s.dispatchEvent(e);}"
            )
        except Exception:
            pass
        # Дождаться появления полей периода
        try:
            self.wait.until(EC.presence_of_element_located((By.ID, 'DatePeriodStart')))
            self.wait.until(EC.presence_of_element_located((By.ID, 'DatePeriodEnd')))
        except Exception:
            pass
        # Установить даты (одинаковые для одного дня)
        try:
            self.driver.execute_script(
                """
                var s=document.getElementById('DatePeriodStart'); var e=document.getElementById('DatePeriodEnd');
                if(s){ s.value=arguments[0]; s.dispatchEvent(new Event('input',{bubbles:true})); s.dispatchEvent(new Event('change',{bubbles:true})); }
                if(e){ e.value=arguments[0]; e.dispatchEvent(new Event('input',{bubbles:true})); e.dispatchEvent(new Event('change',{bubbles:true})); }
                """,
                date_str,
            )
        except Exception:
            pass

        # Сигнатура текущей таблицы до построения
        old_sig = None
        try:
            old_sig = self.driver.execute_script(
                "var t=document.querySelector('table.table.table-nonfluid tbody'); return t ? t.innerText.length : null;"
            )
        except Exception:
            pass

        # Нажать кнопку построения
        clicked = False
        try:
            el = self.wait.until(EC.element_to_be_clickable((By.ID, 'buildReportButton')))
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            except Exception:
                pass
            el.click()
            clicked = True
        except Exception:
            pass

        if not clicked:
            try:
                self.driver.execute_script("if (typeof buildReport === 'function') buildReport();")
                clicked = True
            except Exception:
                clicked = False

        if not clicked:
            for by, sel in [
                (By.XPATH, "//input[@id='buildReportButton' and @value='Построить']"),
                (By.XPATH, "//button[normalize-space()='Построить']"),
                (By.XPATH, "//input[@type='button' and @value='Построить']"),
                (By.CSS_SELECTOR, "#buildReportButton, [name='reportButton']"),
            ]:
                try:
                    el = self.wait.until(EC.element_to_be_clickable((by, sel)))
                    el.click()
                    clicked = True
                    break
                except Exception:
                    continue

        if old_sig is not None:
            for _ in range(200):
                try:
                    new_sig = self.driver.execute_script(
                        "var t=document.querySelector('table.table.table-nonfluid tbody'); return t ? t.innerText.length : null;"
                    )
                    if new_sig != old_sig:
                        break
                except Exception:
                    pass
                time.sleep(0.05)

    def read_table_rows(self) -> List[Tuple[str, List[str]]]:
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table.table-nonfluid tbody tr")))
        rows = self.driver.find_elements(By.CSS_SELECTOR, "table.table.table-nonfluid tbody tr")
        result: List[Tuple[str, List[str]]] = []
        for tr in rows:
            tds = tr.find_elements(By.TAG_NAME, "td")
            if not tds:
                continue
            name = (tds[0].text or "").strip()
            values: List[str] = []
            cols = tds[1:6]
            for td in cols:
                txt = (td.text or "").strip().replace("\xa0", "").replace(" ", "")
                values.append(txt)
                try:
                    self.driver.execute_script(
                        "arguments[0].style.backgroundColor='#00ff00';arguments[0].style.color='#000';",
                        td,
                    )
                except Exception:
                    pass
            result.append((name, values))
        return result

    # ---------- Даты и CSV ----------
    def compute_dates(self) -> List[datetime.date]:
        today = datetime.date.today()
        start = today.replace(day=1)
        yesterday = today - datetime.timedelta(days=1)
        if yesterday < start:
            print("[DATES] Сегодня 1-е: диапазон пуст.")
            return []
        return [start + datetime.timedelta(days=i) for i in range((yesterday - start).days + 1)]

    def _parse_date(self, s: str) -> Optional[datetime.date]:
        s = (s or "").strip()
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

    def reset_csv(self):
        try:
            d = os.path.dirname(self.csv_file)
            if d:
                os.makedirs(d, exist_ok=True)
        except Exception:
            pass
        with open(self.csv_file, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=';')
            w.writerow([
                "Город", "Отдел", "Дата", "Категория",
                "Продажи", "Производство", "Питание персонала", "Отмена", "Брак"
            ])

    def append_csv_rows(self, rows: List[List[str]]):
        if not rows:
            return
        with open(self.csv_file, "a", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=';')
            w.writerows(rows)
        try:
            responses = self._push_supabase_rows(rows)
        except Exception as exc:
            error_text = f"Ошибка отправки: {exc}"
            print(f"[SUPABASE] {error_text}")
            log_bugreport(
                f"MaterialConsumptionReport — ERROR (rows: {len(rows)})",
                0,
                "",
                error_text,
            )
            return
        if not responses:
            return
        for idx, resp in enumerate(responses, start=1):
            context = f"MaterialConsumptionReport — chunk {idx}/{len(responses)} (rows: {len(rows)})"
            log_bugreport(
                context,
                int(resp.get("count", 0)),
                str(resp.get("stdout", "")),
                str(resp.get("stderr", "")),
            )

    # ---------- Интеграция с Supabase ----------
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
        if len(row) < 9:
            return None
        city, dept, date_str, category, sales, production, staff_food, cancel, defect = row[:9]
        report_date = self._safe_parse_date(date_str)
        fm = self.supabase_field_map
        return {
            fm["city"]: city,
            fm["department"]: dept,
            fm["date"]: report_date,
            fm["category"]: category,
            fm["sales"]: self._parse_number(sales),
            fm["production"]: self._parse_number(production),
            fm["staff_food"]: self._parse_number(staff_food),
            fm["cancellation"]: self._parse_number(cancel),
            fm["defect"]: self._parse_number(defect),
        }

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

    # ---------- Вспомогательная диагностика ----------
    def _debug_dump(self, tag: str = "debug"):
        try:
            ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            base_dir = os.path.join("logs", "debug")
            try:
                os.makedirs(base_dir, exist_ok=True)
            except Exception:
                pass
            base = os.path.join(base_dir, f"{ts}-{tag}")
            try:
                self.driver.save_screenshot(base + ".png")
            except Exception:
                pass
            try:
                with open(base + ".html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source or "")
            except Exception:
                pass
            try:
                print(f"[DEBUG] dump saved: {base}.png, {base}.html")
            except Exception:
                pass
        except Exception:
            pass

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
                args = {"userAgent": new_ua or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
                        "platform": "Win32"}
                self.driver.execute_cdp_cmd("Network.setUserAgentOverride", args)
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

    # ---------- Основной сценарий ----------
    def process_city(self, city_name: str, city_uuid: str, dates: List[datetime.date]):
        print("\n" + "#" * 80)
        print(f"[CITY] {city_name}")
        self.select_city(city_uuid)
        self.open_report_for_city(city_uuid)

        departments = self.get_departments()
        print(f"[DEPTS] Найдено отделов: {len(departments)}")

        city_rows: List[List[str]] = []
        for dept in departments:
            print("\n" + "=" * 80)
            print(f"[DEPT] {dept}")
            self.choose_department(dept)
            for dt in dates:
                self.build_for_date(dt)
                rows = self.read_table_rows()
                for cat, vals in rows:
                    city_rows.append([
                        city_name,
                        dept,
                        dt.strftime("%d.%m.%Y"),
                        cat,
                        *vals
                    ])
                print(f"[CSV] {dept} — {dt:%d.%m.%Y}: {len(rows)} строк")
        if city_rows:
            self.append_csv_rows(city_rows)
            print(f"[SUPABASE] {city_name}: выгружено {len(city_rows)} строк.")

    def run(self):
        self.launch_chrome()
        self.connect_driver()
        dates = self.prompt_date_range()
        if dates:
            print(f"[DATES] Диапазон: {dates[0]:%d.%m.%Y} — {dates[-1]:%d.%m.%Y} (всего {len(dates)})")
        else:
            print("[DATES] Диапазон дат пуст — ничего обрабатывать.")
        self.reset_csv()

        cities = self.get_cities()
        print(f"[CITIES] К обработке: {[c[0] for c in cities]}")
        for cidx, (city_name, city_uuid) in enumerate(cities, start=1):
            print(f"[CITY IDX] ({cidx}/{len(cities)})")
            try:
                self.process_city(city_name, city_uuid, dates)
            except Exception as e:
                print(f"[WARN] Ошибка в городе {city_name}: {e}")
            try:
                self.back_to_select_role()
            except Exception as e:
                print(f"[WARN] Не удалось вернуться на SelectRole: {e}")

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
    bot = OfficeMaterialConsumptionReporter()
    try:
        bot.run()
    finally:
        bot.close()
