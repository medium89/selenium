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
CSV_FILE = "отчет.csv"
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
        print("[INIT] Перезапускаю Chrome и настраиваю remote‑debugging…")
        subprocess.run("taskkill /F /IM chrome.exe 2>nul", shell=True)
        chrome_exe = rf"{os.environ['ProgramFiles']}\Google\Chrome\Application\chrome.exe"
        user_dir = os.path.join(os.environ.get("TEMP", os.getcwd()), f"chrome{self.port}")
        subprocess.Popen([
            chrome_exe,
            f"--remote-debugging-port={self.port}",
            f"--user-data-dir={user_dir}",
            "--no-first-run",
            "--no-default-browser-check",
        ])
        if not self._wait_port(self.port, 10):
            raise RuntimeError(f"Порт {self.port} не открылся")

    # Подключение Selenium к уже запущенному Chrome
    def connect_driver(self):
        print("[DRIVER] Подключаюсь к Chrome по debuggerAddress…")
        options = webdriver.ChromeOptions()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{self.port}")
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
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

    # На странице SelectRole нажать роль "Менеджер проектов"
    def ensure_role_selected(self, city_uuid: Optional[str] = None):
        if "/SelectRole" in self.driver.current_url:
            print("[AUTH] Выбираю роль: Менеджер проектов…")
            try:
                self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="roleId"][value="{ROLE_ID}"]'))).click()
            except Exception:
                # Попробуем вариант через input[type=submit]
                try:
                    self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'[name="roleId"][value="{ROLE_ID}"]'))).click()
                except Exception:
                    pass
            # Ждём ухода со страницы выбора роли
            try:
                WebDriverWait(self.driver, 10).until(lambda d: "/SelectRole" not in d.current_url)
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
        self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'button[name="uuid"][value="{city_uuid}"]'))).click()
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
        self.ensure_role_selected()
        # После выбора роли обычно оказываемся на списке подразделений предыдущего города,
        # поэтому перейдём на SelectDepartment для явного выбора следующего города
        self.open_select_department()

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

    # Установка дат и построение отчёта
    def build_for_date(self, dt: datetime.date):
        str_day = str(dt.day)
        for field in ("StartDate", "EndDate"):
            self.wait.until(EC.element_to_be_clickable((By.ID, field))).click()
            self.wait.until(EC.element_to_be_clickable((By.LINK_TEXT, str_day))).click()
            self.click_blank()
        try:
            old_html = self.driver.find_element(By.ID, "report").get_attribute("innerHTML")
        except Exception:
            old_html = None
        self.driver.find_element(By.NAME, "reportButton").click()
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
        today = datetime.date.today()
        start = today.replace(day=1)
        yesterday = today - datetime.timedelta(days=1)
        if yesterday < start:
            dates = []
            print("[DATES] Сегодня 1-е число: диапазон дат пуст (до вчерашнего дня).")
        else:
            days = (yesterday - start).days + 1
            dates = [start + datetime.timedelta(days=i) for i in range(days)]
            print(f"[DATES] Обрабатываю даты: {start:%d.%m.%Y} — {yesterday:%d.%m.%Y} (всего {len(dates)})")

        self.reset_csv()

        # Перебор всех городов с возвратом на SelectDepartment между итерациями
        cities = self.get_cities()
        print(f"[CITIES] К обработке: {[c[0] for c in cities]}")

        for cidx, (city_name, city_uuid) in enumerate(cities, start=1):
            print("\n" + "#" * 80)
            print(f"[CITY] ({cidx}/{len(cities)}) {city_name}")
            try:
                # Вход в город и переход на страницу отчёта
                self.select_city(city_uuid)
                self.open_report_for_city(city_uuid)
                self.select_all_reasons()

                # Список отделов в выбранном городе
                depts = self.get_departments(limit=None)
                print(f"[DEPTS] Отделы: {depts}")

                # Заголовок города в CSV
                self.append_csv_row([f"ГОРОД: {city_name}", ""]) 

                for didx, dept in enumerate(depts, start=1):
                    print("\n" + "=" * 80)
                    print(f"[DEPT] ({didx}/{len(depts)}) {dept}")
                    self.choose_department(dept)
                    self.append_csv_row([f"ОТДЕЛ: {dept}", ""])  # заголовок группы
                    for dt in dates:
                        self.build_for_date(dt)
                        val = self.read_total_value()
                        self.append_csv_row([dt.strftime("%d.%m.%Y"), val])
                        print(f"[CSV] {dt:%d.%m.%Y}: {val}")
                    self.append_csv_row(["", ""])  # разделитель по отделу

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
