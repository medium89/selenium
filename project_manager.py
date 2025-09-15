from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from typing import Optional

import csv
import datetime
import os
import socket
import subprocess
import time


# Конфигурация по умолчанию
PORT = 9222
CSV_FILE = "project.scv"
REPORT_URL = "https://officemanager.dodopizza.ru/OfficeManager/Debiting/PrepareExcelReport"
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ROLE_ID = "8"  # роль, используемая для отчётов
SLOW_DELAY = float(os.environ.get("SLOW_DELAY", "0"))


class DodoDebitingReporter:
    # Инициализация зависимостей и параметров
    def __init__(self, port: int = PORT, csv_file: str = CSV_FILE, url: str = REPORT_URL, slow: float = SLOW_DELAY):
        self.port = port
        self.csv_file = csv_file
        self.url = url
        self.slow = slow
        self.driver = None
        self.wait = None

    # Запуск Chrome с remote-debugging и ожиданием готовности порта
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

    # Подключение Selenium к уже запущенному Chrome
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
            # Постоянный профиль (вариант C)
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

    # Открытие страницы выбора города (SelectDepartment)
    def open_select_department(self):
        print("[NAV] Перехожу на экран выбора города…")
        self.driver.get(SELECT_DEPARTMENT_URL)
        # Если редиректнуло на выбор роли, выберем Менеджер проектов
        self.ensure_role_selected()
        # После выбора роли повторно перейдём на SelectDepartment при необходимости
        if "/SelectDepartment" not in self.driver.current_url:
            try:
                self.driver.get(SELECT_DEPARTMENT_URL)
            except Exception:
                pass

    # Клик по роли "Менеджер проектов" на SelectRole
    def choose_role(self):
        print("[AUTH] Выбираю роль: Менеджер проектов…")
        try:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="roleId"][value="{ROLE_ID}"]'))).click()
        except Exception:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'[name="roleId"][value="{ROLE_ID}"]'))).click()
        # Ждём ухода со страницы выбора роли
        try:
            WebDriverWait(self.driver, 10).until(lambda d: "/SelectRole" not in d.current_url)
        except Exception:
            pass

    # На странице SelectRole нажать роль; при переданном city_uuid попробовать кликнуть город
    def ensure_role_selected(self, city_uuid: Optional[str] = None):
        if "/SelectRole" in self.driver.current_url:
            self.choose_role()
            if city_uuid:
                # Если после выбора роли осталось на той же странице или сразу видны варианты города
                try:
                    self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="uuid"][value="{city_uuid}"]'))).click()
                except Exception:
                    pass

    # Получение списка городов на странице SelectDepartment: [(name, uuid), ...]
    def get_cities(self):
        print("[CITIES] Собираю список городов…")
        self.open_select_department()
        # Подождём появления любых кнопок выбора города
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
        # Уникализируем по uuid и сортируем по имени
        seen = set()
        cities = []
        for it in items:
            uuid = it.get('uuid')
            name = it.get('name')
            if uuid and uuid not in seen:
                seen.add(uuid)
                cities.append((name, uuid))
        cities.sort(key=lambda x: x[0].lower())
        if not cities:
            raise RuntimeError("Не удалось получить список городов на SelectDepartment")
        print(f"[CITIES] Найдено городов: {len(cities)}")
        return cities

    # Выбрать город по uuid на SelectDepartment
    def select_city(self, city_uuid: str):
        self.open_select_department()
        # На всякий случай, если всё ещё попали на SelectRole, выберем роль
        self.ensure_role_selected()
        # Клик по кнопке/ссылке города
        try:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="uuid"][value="{city_uuid}"]'))).click()
        except Exception:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'a[name="uuid"][value="{city_uuid}"]'))).click()
        time.sleep(0.2)

    # Открыть страницу отчёта, обработать возможный редирект через SelectRole
    def open_report_for_city(self, city_uuid: str):
        print("[NAV] Перехожу на страницу отчётов…")
        self.driver.get(REPORT_URL)
        self.ensure_role_selected(city_uuid)

    # Вернуться на SelectRole и выбрать роль
    def back_to_select_role(self):
        print("[NAV] Возвращаюсь на SelectRole для смены контекста…")
        try:
            self.driver.get(BACK_TO_SELECT_ROLE_URL)
        except Exception:
            pass
        # Дождёмся появления/редиректа на SelectRole и выберем роль
        try:
            WebDriverWait(self.driver, 10).until(EC.url_contains("/SelectRole"))
        except Exception:
            pass
        if "/SelectRole" in self.driver.current_url:
            self.choose_role()
        # После выбора роли обычно оказываемся на списке подразделений предыдущего города,
        # поэтому перейдём на SelectDepartment для явного выбора следующего города
        self.open_select_department()

    # Подсчёт дат: с 1-го числа текущего месяца до вчера
    def compute_dates(self) -> list:
        today = datetime.date.today()
        start = today.replace(day=1)
        yesterday = today - datetime.timedelta(days=1)
        if yesterday < start:
            print("[DATES] Сегодня 1-е число: диапазон дат пуст (до вчерашнего дня).")
            return []
        days = (yesterday - start).days + 1
        dates = [start + datetime.timedelta(days=i) for i in range(days)]
        print(f"[DATES] Обрабатываю даты: {start:%d.%m.%Y} — {yesterday:%d.%m.%Y} (всего {len(dates)})")
        return dates

    # Обработка одного города целиком
    def process_city(self, city_name: str, city_uuid: str, dates: list):
        print("\n" + "#" * 80)
        print(f"[CITY] {city_name}")
        # Вход в город и переход на страницу отчёта
        self.select_city(city_uuid)
        self.open_report_for_city(city_uuid)
        self.select_all_reasons()

        # Список отделов в выбранном городе
        depts = self.get_departments(limit=None)
        print(f"[DEPTS] Отделы: {depts}")

        # Заголовок города в CSV
        self.append_csv_row([f"ГОРОД: {city_name}", ""]) 

        for dept in depts:
            print("\n" + "=" * 80)
            print(f"[DEPT] {dept}")
            self.choose_department(dept)
            self.append_csv_row([f"ОТДЕЛ: {dept}", ""])  # заголовок группы
            for dt in dates:
                self.build_for_date(dt)
                val = self.read_total_value()
                self.append_csv_row([dt.strftime("%d.%m.%Y"), val])
                print(f"[CSV] {dt:%d.%m.%Y}: {val}")
            self.append_csv_row(["", ""])  # разделитель по отделу

    # Нажатие в пустую область страницы (сбрасывает фокус/оверлеи)
    def click_blank(self):
        try:
            self.driver.find_element(By.CSS_SELECTOR, ".content").click()
        except Exception:
            try:
                self.driver.find_element(By.TAG_NAME, "body").click()
            except Exception:
                pass

    # Показ/скрытие красного баннера с предупреждением
    def show_red_banner(self, message: str):
        js = """
        (function(msg){
            var id='selenium-warn-banner', el=document.getElementById(id);
            if(!el){ el=document.createElement('div'); el.id=id; el.style.cssText='position:fixed;top:0;left:0;right:0;z-index:999999;padding:10px 16px;background:#d9363e;color:#fff;font:600 14px/1 sans-serif;box-shadow:0 2px 6px rgba(0,0,0,.2)'; document.body.appendChild(el);} el.textContent=msg;
        })(arguments[0]);
        """
        try:
            self.driver.execute_script(js, message)
        except Exception:
            pass

    def hide_red_banner(self):
        try:
            self.driver.execute_script("var el=document.getElementById('selenium-warn-banner'); if(el){el.remove();}")
        except Exception:
            pass

    # Выбор всех причин списания через прямую установку в select
    def select_all_reasons(self):
        print("[FILTER] Отмечаю все причины списания…")
        try:
            self.driver.execute_script(
                "var s=document.getElementById('DebitingReasonId'); if(!s) return; Array.from(s.options).forEach(o=>o.selected=true); var e; try{e=new Event('change',{bubbles:true});}catch(err){e=document.createEvent('HTMLEvents'); e.initEvent('change',true,false);} s.dispatchEvent(e);"
            )
        except Exception:
            pass

    # Получение списка отделов из select или раскрытого списка
    def get_departments(self, limit: Optional[int] = None):
        print("[DEPTS] Получаю список отделов…")
        names = []
        for _ in range(100):
            try:
                names = self.driver.execute_script(
                    "return Array.from(document.querySelectorAll('#UnitId option')).map(o=>(o.text||'').trim()).filter(t=>t && t.toLowerCase()!=='выбрать все');"
                ) or []
            except Exception:
                names = []
            if names:
                break
            time.sleep(0.1)
        if not names:
            try:
                opened = self.driver.execute_script(
                    "var s=document.getElementById('UnitId'); if(!s) return false; var box=s.closest('.select-report'); if(!box) return false; var cap=box.querySelector('.CaptionCont'); if(!cap) return false; cap.click(); return true;"
                )
                if opened:
                    time.sleep(0.3)
                    tmp = []
                    for li in self.driver.find_elements(By.CSS_SELECTOR, ".open li"):
                        t = (li.text or "").strip()
                        if t and t.lower() != "выбрать все":
                            tmp.append(t)
                    if tmp:
                        names = tmp
                    self.click_blank()
            except Exception:
                pass
        if not names:
            raise RuntimeError("Список отделов пуст")
        return names if limit is None else names[:limit]

    # Принудительный выбор только одного отдела на уровне select
    def choose_department(self, name: str):
        self.driver.execute_script(
            "var s=document.getElementById('UnitId'); if(!s) return; for(var i=0;i<s.options.length;i++){var o=s.options[i]; o.selected=((o.text||'').trim()===arguments[0]);} var e; try{e=new Event('change',{bubbles:true});}catch(err){e=document.createEvent('HTMLEvents'); e.initEvent('change',true,false);} s.dispatchEvent(e);",
            name,
        )
        if self.slow:
            time.sleep(self.slow)
        try:
            selected = self.driver.execute_script(
                "var s=document.getElementById('UnitId'); if(!s) return []; return Array.from(s.options).filter(o=>o.selected).map(o=>(o.text||'').trim());"
            ) or []
        except Exception:
            selected = []
        if name not in selected or len(selected) != 1:
            self.show_red_banner("Внимание: выбран не один отдел")
        else:
            self.hide_red_banner()

    # Установка дат и построение отчёта (ускоренный ввод как в OfficeManager)
    def build_for_date(self, dt: datetime.date):
        date_str = dt.strftime("%d.%m.%Y")
        # Убедиться, что поля существуют
        try:
            self.wait.until(EC.presence_of_element_located((By.ID, 'StartDate')))
            self.wait.until(EC.presence_of_element_located((By.ID, 'EndDate')))
        except Exception:
            pass
        # Прямое выставление дат через JS с генерацией input/change
        try:
            self.driver.execute_script(
                """
                var s=document.getElementById('StartDate'); var e=document.getElementById('EndDate');
                if(s){ s.value=arguments[0]; s.dispatchEvent(new Event('input',{bubbles:true})); s.dispatchEvent(new Event('change',{bubbles:true})); }
                if(e){ e.value=arguments[0]; e.dispatchEvent(new Event('input',{bubbles:true})); e.dispatchEvent(new Event('change',{bubbles:true})); }
                """,
                date_str,
            )
        except Exception:
            pass

        # Зафиксировать текущий HTML, чтобы дождаться обновления
        try:
            old_html = self.driver.find_element(By.ID, "report").get_attribute("innerHTML")
        except Exception:
            old_html = None

        # Нажать кнопку формирования отчёта
        try:
            btn = self.wait.until(EC.element_to_be_clickable((By.NAME, "reportButton")))
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            except Exception:
                pass
            btn.click()
        except Exception:
            # Резерв на случай альтернативной разметки
            try:
                self.driver.find_element(By.CSS_SELECTOR, "button[name='reportButton'], input[name='reportButton']").click()
            except Exception:
                pass

        if old_html is not None:
            for _ in range(200):
                try:
                    if self.driver.find_element(By.ID, "report").get_attribute("innerHTML") != old_html:
                        break
                except Exception:
                    pass
                time.sleep(0.05)

    # Извлечение итогового значения из таблицы отчёта
    def read_total_value(self) -> str:
        cells = self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody td.totalValue")))
        non_empty = [td for td in cells if td.text.strip() and any(c.isdigit() for c in td.text)]
        if not non_empty:
            return "0"
        target = non_empty[-1]
        try:
            self.driver.execute_script("arguments[0].style.backgroundColor='#00ff00';arguments[0].style.color='#000';", target)
        except Exception:
            pass
        return target.text.replace("\xa0", "").replace(" ", "").replace("₽", "")

    # Сброс/создание CSV и дописывание строк
    def reset_csv(self):
        # Создаём файл в кодировке UTF‑8 с BOM, чтобы Excel корректно показал кириллицу
        with open(self.csv_file, "w", encoding="utf-8-sig", newline="") as f:
            f.write("\ufeff")

    def append_csv_row(self, row):
        # Дозапись строк в UTF‑8 с BOM (BOM уже на первом открытии)
        with open(self.csv_file, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter=";").writerow(row)

    # Основной сценарий выполнения
    def run(self):
        self.launch_chrome()
        self.connect_driver()

        # Даты для обработки: с 1-го числа по вчерашний день
        dates = self.compute_dates()

        self.reset_csv()

        # Перебор всех городов с возвратом на SelectDepartment между итерациями
        cities = self.get_cities()
        print(f"[CITIES] К обработке: {[c[0] for c in cities]}")

        for cidx, (city_name, city_uuid) in enumerate(cities, start=1):
            print(f"[CITY IDX] ({cidx}/{len(cities)})")
            try:
                self.process_city(city_name, city_uuid, dates)
            except Exception as e:
                print(f"[WARN] Ошибка при обработке города {city_name}: {e}")
                self.append_csv_row([f"ГОРОД: {city_name}", f"ОШИБКА: {e}"])
                self.append_csv_row(["", ""])  # разделитель

            # После завершения города переходим на BackToSelectRole и выбираем роль
            try:
                self.back_to_select_role()
            except Exception as e:
                print(f"[WARN] Не удалось вернуться на SelectRole: {e}")

        print(f"[DONE] Готово! Файл {self.csv_file} сохранён.")

    # Корректное завершение драйвера
    def close(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass

    # Ожидание открытия TCP‑порта
    @staticmethod
    def _wait_port(port: int, timeout: int = 10) -> bool:
        for _ in range(timeout * 10):
            with socket.socket() as s:
                if s.connect_ex(("127.0.0.1", port)) == 0:
                    return True
            time.sleep(0.1)
        return False


if __name__ == "__main__":
    bot = DodoDebitingReporter()
    try:
        bot.run()
    finally:
        bot.close()
