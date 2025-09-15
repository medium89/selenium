from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from typing import Optional, List, Tuple

import csv
import datetime
import os
import socket
import subprocess
import time


# Конфигурация по умолчанию (страницы и роль зашиты в коде)
PORT = 9222
CSV_FILE = "reports/office.csv"
REPORT_URL = "https://officemanager.dodopizza.ru/OfficeManager/MaterialConsumption"
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ROLE_ID = "7"  # роль Офис‑менеджера
SLOW_DELAY = float(os.environ.get("SLOW_DELAY", "0"))


class OfficeMaterialConsumptionReporter:
    """Сбор данных по небольшому отчёту MaterialConsumption.

    Структурно повторяет project_manager.py: страницы и роль зашиты,
    последовательность переходов идентичная.
    """

    def __init__(self, port: int = PORT, csv_file: str = CSV_FILE, url: str = REPORT_URL, slow: float = SLOW_DELAY):
        self.port = port
        self.csv_file = csv_file
        self.url = url
        self.slow = slow
        self.driver: Optional[webdriver.Chrome] = None
        self.wait: Optional[WebDriverWait] = None

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
                user_dir = os.environ.get("USER_DATA_DIR") or os.path.join(os.environ.get("TEMP", os.getcwd()), f"chrome{self.port}")
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

    def _make_service(self) -> Service:
        path = os.environ.get("CHROMEDRIVER", "/usr/bin/chromedriver")
        if path and os.path.exists(path):
            return Service(path)
        return Service(ChromeDriverManager().install())

    def connect_driver(self):
        print("[DRIVER] Инициализация драйвера Chrome…")
        options = webdriver.ChromeOptions()
        if self._wait_port(self.port, 1):
            print("[DRIVER] Найден debuggerAddress — подключаюсь к внешнему Chrome…")
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{self.port}")
            self.driver = webdriver.Chrome(service=self._make_service(), options=options)
        else:
            if os.environ.get("CHROME_BIN"):
                options.binary_location = os.environ["CHROME_BIN"]
            user_dir = os.environ.get("USER_DATA_DIR")
            if user_dir:
                try:
                    os.makedirs(user_dir, exist_ok=True)
                except Exception:
                    pass
                options.add_argument(f"--user-data-dir={user_dir}")
            if os.environ.get("HEADLESS", "0") == "1":
                options.add_argument("--headless=new")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
            self.driver = webdriver.Chrome(service=self._make_service(), options=options)
        self.wait = WebDriverWait(self.driver, 25)

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

    # ---------- Города ----------
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
            cols = tds[1:6]  # первые 5 числовых колонок
            for td in cols:
                txt = (td.text or "").strip().replace("\xa0", "").replace(" ", "")
                values.append(txt)
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

    def reset_csv(self):
        # Ensure target directory exists
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
        with open(self.csv_file, "a", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=';').writerows(rows)

    # ---------- Основной сценарий ----------
    def process_city(self, city_name: str, city_uuid: str, dates: List[datetime.date]):
        print("\n" + "#" * 80)
        print(f"[CITY] {city_name}")
        self.select_city(city_uuid)
        self.open_report_for_city(city_uuid)

        departments = self.get_departments()
        print(f"[DEPTS] Найдено отделов: {len(departments)}")

        for dept in departments:
            print("\n" + "=" * 80)
            print(f"[DEPT] {dept}")
            self.choose_department(dept)
            for dt in dates:
                self.build_for_date(dt)
                rows = self.read_table_rows()
                out_rows: List[List[str]] = []
                for cat, vals in rows:
                    out_rows.append([
                        city_name,
                        dept,
                        dt.strftime("%d.%m.%Y"),
                        cat,
                        *vals
                    ])
                self.append_csv_rows(out_rows)
                print(f"[CSV] {dept} — {dt:%d.%m.%Y}: {len(rows)} строк")

    def run(self):
        self.launch_chrome()
        self.connect_driver()
        dates = self.compute_dates()
        self.reset_csv()

        cities = self.get_cities()
        print(f"[CITIES] К обработке: {[c[0] for c in cities]}")
        for cidx, (city_name, city_uuid) in enumerate(cities, start=1):
            print(f"[CITY IDX] ({cidx}/{len(cities)})")
            try:
                self.process_city(city_name, city_uuid, dates)
            except Exception as e:
                print(f"[WARN] Ошибка в городе {city_name}: {e}")
                # Продолжим со следующими городами
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
