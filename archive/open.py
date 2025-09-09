from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import csv
from datetime import datetime
from zoneinfo import ZoneInfo
import subprocess
import time
import socket
import os
import json

URL = "https://officemanager.dodopizza.ru/OfficeManager/Debiting/PrepareExcelReport"

CHROME_PATHS = [
    r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
]
USER_DATA_DIR = r"C:\\Users\\ivano\\AppData\\Local\\Google\\Chrome\\User Data"
PROFILE_DIR = "Default"  # если основной профиль другой: "Profile 1", "Profile 2", ...
DEBUG_PORT = 9222


def is_port_open(host: str, port: int, timeout: float = 0.5) -> bool:
    s = socket.socket()
    s.settimeout(timeout)
    try:
        s.connect((host, port))
        return True
    except Exception:
        return False
    finally:
        try:
            s.close()
        except Exception:
            pass


def ensure_chrome_debug_session(initial_url: str | None = None) -> None:
    # Если порт уже открыт, предполагаем, что Chrome запущен с remote debugging
    if is_port_open("127.0.0.1", DEBUG_PORT):
        return

    # Ищем chrome.exe
    chrome_exe = next((p for p in CHROME_PATHS if os.path.exists(p)), None)
    if not chrome_exe:
        raise FileNotFoundError("Не найден chrome.exe. Проверь путь в CHROME_PATHS.")

    # Важно: для использования основного профиля надо закрыть все окна Chrome
    # Иначе запустить второй экземпляр с тем же профилем не получится
    launch_cmd = [
        chrome_exe,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={USER_DATA_DIR}",
        f"--profile-directory={PROFILE_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
    ]

    # Если задан стартовый адрес — открываем его сразу
    if initial_url:
        launch_cmd.append(initial_url)

    # Запускаем обычный Chrome (НЕ через WebDriver)
    subprocess.Popen(launch_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # Ждем, пока откроется порт отладки
    for _ in range(50):  # ~5 секунд
        if is_port_open("127.0.0.1", DEBUG_PORT):
            break
        time.sleep(0.1)
    else:
        raise RuntimeError("Chrome не запустился с remote debugging. Закрой Chrome и попробуй снова.")


# ==== Helpers from test_1 (адаптированы) ====
def safe_click(driver, locator, timeout=20):
    """Ищем кликабельный элемент в документе или во фреймах (до 2 уровней) и кликаем."""
    def try_in_current_context() -> bool:
        el = WebDriverWait(driver, max(3, timeout // 3)).until(EC.element_to_be_clickable(locator))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        try:
            el.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", el)
        return True

    last_err = None
    # 0) По умолчанию
    try:
        driver.switch_to.default_content()
        if try_in_current_context():
            return
    except Exception as e:
        last_err = e

    # 1) Топ-уровень фреймов
    top_frames = driver.find_elements(By.TAG_NAME, "iframe")
    for i in range(len(top_frames)):
        try:
            driver.switch_to.default_content()
            # получаем фрейм по индексу повторно (избегаем stale)
            f = driver.find_elements(By.TAG_NAME, "iframe")[i]
            driver.switch_to.frame(f)
            if try_in_current_context():
                return
        except Exception as e:
            last_err = e
        # 2) Вложенные фреймы уровня 2
        try:
            inner_frames = driver.find_elements(By.TAG_NAME, "iframe")
            for j in range(len(inner_frames)):
                try:
                    driver.switch_to.default_content()
                    f1 = driver.find_elements(By.TAG_NAME, "iframe")[i]
                    driver.switch_to.frame(f1)
                    f2 = driver.find_elements(By.TAG_NAME, "iframe")[j]
                    driver.switch_to.frame(f2)
                    if try_in_current_context():
                        return
                except Exception as e2:
                    last_err = e2
                    continue
        except Exception as e3:
            last_err = e3
            continue

    driver.switch_to.default_content()
    raise TimeoutException(f"Не удалось кликнуть по {locator}. Последняя ошибка: {last_err}")


def get_berlin_date_iso():
    return datetime.now(ZoneInfo("Europe/Berlin")).date().isoformat()


def clean_money(text: str) -> str:
    s = text.replace("\xa0", " ").replace("₽", "").strip()
    return s


def find_elements_anywhere(driver, locator):
    found = []
    try:
        driver.switch_to.default_content()
        found.extend(driver.find_elements(*locator))
    except Exception:
        pass
    # top-level frames
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for i in range(len(frames)):
        try:
            driver.switch_to.default_content()
            f1 = driver.find_elements(By.TAG_NAME, "iframe")[i]
            driver.switch_to.frame(f1)
            found.extend(driver.find_elements(*locator))
        except Exception:
            pass
        # nested frames
        try:
            inner_frames = driver.find_elements(By.TAG_NAME, "iframe")
            for j in range(len(inner_frames)):
                try:
                    driver.switch_to.default_content()
                    f1 = driver.find_elements(By.TAG_NAME, "iframe")[i]
                    driver.switch_to.frame(f1)
                    f2 = driver.find_elements(By.TAG_NAME, "iframe")[j]
                    driver.switch_to.frame(f2)
                    found.extend(driver.find_elements(*locator))
                except Exception:
                    pass
        except Exception:
            pass
    driver.switch_to.default_content()
    return found


def wait_presence_anywhere(driver, locator, timeout=30):
    end = time.time() + timeout
    while time.time() < end:
        els = find_elements_anywhere(driver, locator)
        if els:
            return els
        time.sleep(0.2)
    raise TimeoutException(f"Элемент {locator} не найден ни в одном фрейме за {timeout}с")


def click_date_by_day(driver, input_id: str, day: int, timeout_show=10, timeout_click=10):
    """Открывает jQuery UI datepicker кликом по полю и кликает по дню."""
    # Клик по полю, чтобы открыть календарь
    safe_click(driver, (By.ID, input_id), timeout=max(5, timeout_show))
    # Иногда удобнее подсказать показать календарь явно
    try:
        driver.execute_script("if (window.jQuery && jQuery('#%s').datepicker) jQuery('#%s').datepicker('show');" % (input_id, input_id))
    except Exception:
        pass
    # Ждём контейнер календаря
    WebDriverWait(driver, timeout_show).until(
        EC.visibility_of_element_located((By.ID, "ui-datepicker-div"))
    )
    # Кликаем по нужному дню
    day_xpath = f"//div[@id='ui-datepicker-div']//table[contains(@class,'ui-datepicker-calendar')]//a[normalize-space(text())='{day}']"
    safe_click(driver, (By.XPATH, day_xpath), timeout=timeout_click)


# 1) Гарантируем запущенный Chrome с твоим профилем и портом отладки
ensure_chrome_debug_session(URL)

# 2) Подключаемся к уже запущенному Chrome через DevTools
options = Options()
options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")

# Убираем часть автоматизационных флагов, чтобы не мешали
options.add_experimental_option("excludeSwitches", ["enable-automation"])  # скрыть баннер
options.add_experimental_option("useAutomationExtension", False)

# Логи браузера и производительности (для сетевых статусов)
options.set_capability('goog:loggingPrefs', {
    'browser': 'ALL',
    'performance': 'ALL',
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 30)

def log(msg: str):
    print(f"[open] {msg}")

def dump_console_and_network(prefix: str = ""):
    # Console logs
    try:
        logs = driver.get_log('browser')
        if logs:
            print(f"[open] {prefix} Консоль браузера ({len(logs)}):")
            for entry in logs[-50:]:
                level = entry.get('level')
                message = entry.get('message', '')
                ts = entry.get('timestamp')
                if level in ('SEVERE', 'ERROR', 'WARNING'):
                    print(f"[open] [{level}] {message}")
    except Exception:
        pass

    # Performance logs: ищем Network.* c ошибками/404
    bad_urls = []
    try:
        perflogs = driver.get_log('performance')
        for pl in perflogs[-200:]:
            try:
                msg = json.loads(pl.get('message', '{}'))
                m = msg.get('message', {})
                method = m.get('method')
                params = m.get('params', {})
                if method == 'Network.responseReceived':
                    resp = params.get('response', {})
                    status = int(resp.get('status', 0))
                    url = resp.get('url', '')
                    if status >= 400 and url:
                        bad_urls.append((status, url))
                elif method == 'Network.loadingFailed':
                    url = params.get('requestId', '')
                    error_text = params.get('errorText', '')
                    bad_urls.append((0, f"LoadingFailed {error_text} id={url}"))
            except Exception:
                continue
    except Exception:
        perflogs = []

    if bad_urls:
        print("[open] Сетевые ошибки/статусы:")
        for status, url in bad_urls[:20]:
            print(f"[open] HTTP {status}: {url}")

# Всегда открываем новую вкладку и идём на нужный URL, чтобы точно управлять активной вкладкой
log("Открываю новую вкладку с целевым URL...")
driver.switch_to.new_window('tab')
driver.get(URL)
wait.until(EC.url_contains("OfficeManager/Debiting/PrepareExcelReport"))
dump_console_and_network("После открытия URL:")

# Закрываем все другие вкладки, оставляем только текущую
if len(driver.window_handles) > 1:
    good = driver.current_window_handle
    for h in list(driver.window_handles):
        if h != good:
            driver.switch_to.window(h)
            driver.close()
    driver.switch_to.window(good)

# ==== Дальнейшие действия со страницы (без кликов) ====
try:
    wait_presence_anywhere(driver, (By.ID, "StartDate"), timeout=20)
except TimeoutException:
    log("Не вижу поля StartDate. Возможно, требуется авторизация. Выполни вход вручную.")
    input("После входа нажми Enter, я продолжу...")
    driver.get(URL)
    wait_presence_anywhere(driver, (By.ID, "StartDate"), timeout=30)
    dump_console_and_network("После авторизации и возврата на страницу:")

log("Выбираю даты кликами по календарю…")
now = datetime.now()
start_day = 1
end_day = 5
clicked_dates = False
try:
    click_date_by_day(driver, "StartDate", start_day)
    time.sleep(0.2)
    click_date_by_day(driver, "EndDate", end_day)
    clicked_dates = True
    log("Даты выбраны через календарь.")
except Exception as e:
    log(f"Не получилось выбрать даты кликами: {e}. Пытаюсь установить через JS…")
    js_set_dates = """
    return (function(){
      function pad2(n){return (n<10?'0':'')+n;}
      function fmt(y,m,d,fmt){
        const dd=pad2(d), mm=pad2(m);
        const yyyy=String(y), yy=String(y%100).padStart(2,'0');
        if(!fmt) fmt = (document.getElementById('datePickerDateFormat')||{}).value || 'dd.mm.yy';
        return fmt
          .replace(/dd/i, dd)
          .replace(/mm/i, mm)
          .replace(/yyyy/i, yyyy)
          .replace(/yy/i, yy);
      }
      function setInput(id, y,m,d){
        const el = document.getElementById(id);
        if(!el) return false;
        try{
          const value = fmt(y,m,d,(document.getElementById('datePickerDateFormat')||{}).value);
          if (window.jQuery){
            jQuery(el).val(value).trigger('input').trigger('change');
          } else {
            el.value = value;
            el.dispatchEvent(new Event('input', {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
          }
          return true;
        }catch(e){ return false; }
      }
      const ok1 = setInput('StartDate', %Y%, %M%, %SD%);
      const ok2 = setInput('EndDate', %Y%, %M%, %ED%);
      return ok1 && ok2;
    })();
    """.replace('%Y%', str(now.year)).replace('%M%', str(now.month)).replace('%SD%', str(start_day)).replace('%ED%', str(end_day))
    ok_dates = driver.execute_script(js_set_dates)
    log(f"Даты установлены через JS: {ok_dates}")

log("Устанавливаю фильтры через JS по id селектов...")
js_set_filters = """
return (function(){
  function dispatch(el){ if(!el) return; el.dispatchEvent(new Event('input',{bubbles:true})); el.dispatchEvent(new Event('change',{bubbles:true})); }
  function selectAllById(id){
    const sel = document.getElementById(id);
    if(!sel || !sel.options) return false;
    let changed=false;
    for(const opt of Array.from(sel.options)){
      if(!opt.disabled){ opt.selected = true; changed=true; }
    }
    if(changed) dispatch(sel);
    return changed;
  }
  const r1 = selectAllById('DebitingReasonId');
  const r2 = selectAllById('UnitId');
  return r1 || r2;
})();
"""
ok_filters = driver.execute_script(js_set_filters)
log(f"Фильтры установлены: {ok_filters}")

log("Нажимаю кнопку построения отчёта…")
try:
    safe_click(driver, (By.NAME, "reportButton"))
except Exception as e:
    log(f"Не удалось нажать кнопку кликом: {e}. Вызываю buildReport()…")
    driver.execute_script("if (typeof buildReport==='function'){ buildReport(); }")
time.sleep(0.5)
dump_console_and_network("После попытки запуска отчёта:")

# Если отчёт не начал строиться (нет контента), пробуем прямой POST через fetch
try:
    has_any = driver.execute_script("return !!document.querySelector('#report table, #productsDebitingReport table');")
except Exception:
    has_any = False
if not has_any:
    log("Пробую построить отчёт прямым POST (обходя onclick)…")
    js_fetch = r"""
    var cb = arguments[arguments.length - 1];
    (async function(){
      try {
        var form = document.getElementById('PrepareExcelReportForm');
        if(!form){ cb('no-form'); return; }
        var params = new URLSearchParams(new FormData(form));
        async function postHtml(url, sel){
          const resp = await fetch(url, {
            method: 'POST',
            headers: {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'},
            body: params.toString(),
            credentials: 'same-origin'
          });
          const html = await resp.text();
          const target = document.querySelector(sel);
          if (target) target.innerHTML = html;
          return {url: url, status: resp.status, len: html.length};
        }
        const r = await Promise.all([
          postHtml('/OfficeManager/Debiting/BuildReport', '#report'),
          postHtml('/OfficeManager/Debiting/BuildProductsDebitingReport', '#productsDebitingReport')
        ]);
        cb(JSON.stringify(r));
      } catch(e){ cb('error:'+ (e && e.message ? e.message : e)); }
    })();
    """
    try:
        res = driver.execute_async_script(js_fetch)
        log(f"Результат fetch: {res}")
    except Exception as e:
        log(f"Ошибка при fetch-построении: {e}")

log("Жду итоговые значения...")
# Ждём, пока в контейнерах отчёта появятся таблицы/данные
wait_presence_anywhere(driver, (By.CSS_SELECTOR, "#report table, #productsDebitingReport table, td.totalValue.text-right.number"), timeout=90)
dump_console_and_network("После появления отчёта:")

# Собираем итоговые ячейки и берём последнюю
cells = find_elements_anywhere(driver, (By.CSS_SELECTOR, "td.totalValue.text-right.number"))
assert cells, "Не найдены ячейки с классом 'totalValue text-right number'."
last_cell = cells[-1]
value_raw = last_cell.text
value_clean = clean_money(value_raw)

line = f"{get_berlin_date_iso()}: {value_clean}"
csv_path = "report_values.csv"
with open(csv_path, "a", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow([line])

log(f"Сохранено: {line} -> {csv_path}")

input("Нажми Enter, чтобы закрыть драйвер (Chrome останется открыт)...")

driver.quit()
