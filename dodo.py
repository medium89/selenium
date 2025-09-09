from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from datetime import date

import traceback   # добавьте в импорты
import subprocess, socket, time, os, csv, datetime

PORT = 9222
CSV_FILE = "отчет.csv"
URL  = "https://officemanager.dodopizza.ru/OfficeManager/Debiting/PrepareExcelReport"

# ---------- 1. Запуск Chrome с remote-debugging ----------
subprocess.run("taskkill /F /IM chrome.exe 2>nul", shell=True)
chrome_exe = rf"{os.environ['ProgramFiles']}\Google\Chrome\Application\chrome.exe"
user_dir   = os.path.join(os.environ["TEMP"], "chrome9222")
subprocess.Popen([
    chrome_exe,
    f"--remote-debugging-port={PORT}",
    f"--user-data-dir={user_dir}",
    "--no-first-run", "--no-default-browser-check"
])

def port_open(p, timeout=10):
    for _ in range(timeout * 10):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) == 0:
                return True
        time.sleep(0.1)
    return False

if not port_open(PORT):
    raise RuntimeError("Порт 9222 не открылся")

# ---------- 2. Подключаемся ----------
options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", f"127.0.0.1:{PORT}")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)

# ---------- 3. Открываем целевую страницу ----------
driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[-1])
driver.get(URL)

# ---------- 4. Выбор роли и города (если попадётся) ----------
if "/SelectRole" in driver.current_url:
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[name="roleId"][value="8"]'))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[name="uuid"][value="000d3a2480c380e711e75b09d9d47a73"]'))).click()
    wait.until(EC.url_contains("/PrepareExcelReport"))

# ---------- 5. Диапазон дат: 1-е число месяца → вчера ----------
today = datetime.date.today()
first = today.replace(day=1)
yesterday = today - datetime.timedelta(days=1)
date_list = [first + datetime.timedelta(days=i) for i in range((yesterday - first).days + 1)]

# ---------- 6. Цикл по датам ----------
with open(CSV_FILE, "w", newline='', encoding='utf-8') as f:
    writer = csv.writer(f, delimiter=':')
    for dt in date_list:
        str_date = dt.strftime("%d.%m.%Y")

        # ------ StartDate = EndDate = текущая дата ------
        for field in ("StartDate", "EndDate"):
            wait.until(EC.element_to_be_clickable((By.ID, field))).click()
            wait.until(EC.element_to_be_clickable((By.LINK_TEXT, str(dt.day)))).click()
            driver.find_element(By.CSS_SELECTOR, ".content").click()

        # ---------- «Выбрать все» в кастом-multiselect ----------
        # 1. Причина списания (вторая колонка)
        driver.find_element(By.CSS_SELECTOR, ".content").click()  # закрыть всё
        driver.find_element(By.CSS_SELECTOR, ".col-md-3:nth-child(1) .CaptionCont i").click()  # открыть список
        wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, ".open li:nth-child(1) i"))).click() # «Выбрать все»
        time.sleep(0.2)  # пусть успеет отметиться

        # 2. Отдел (третья колонка)
        driver.find_element(By.CSS_SELECTOR, ".content").click()  # закрыть
        driver.find_element(By.CSS_SELECTOR, ".col-md-3:nth-child(2) .CaptionCont i").click()  # открыть список
        wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, ".open li:nth-child(1) i"))).click()  # «Выбрать все»
        time.sleep(0.2)




        # ------ формируем отчёт ------
        driver.find_element(By.NAME, "reportButton").click()




        # ------ парсим последнюю НЕПУСТУЮ ячейку totalValue ------
        all_totals = wait.until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "tbody td.totalValue")))   # все ячейки под классом

        # фильтруем только те, у которых есть цифры
        non_empty = [td for td in all_totals if td.text.strip() and any(c.isdigit() for c in td.text)]

        if not non_empty:
            value = "0"
        else:
            target_cell = non_empty[-1]          # последняя непустая
            value = target_cell.text.replace("\xa0", "").replace(" ", "").replace("₽", "")

            # ПОДСВЕТКА
            driver.execute_script(
                "arguments[0].style.backgroundColor = '#00ff00'; arguments[0].style.color = '#000';",
                target_cell)

        writer.writerow([str_date, value])
        print(f"{str_date} : {value}")


        

print(f"Готово! Файл {CSV_FILE} сохранён.")
input('Enter — закрыть драйвер (Chrome останется жив)\n')
driver.quit()