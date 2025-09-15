import os
import sys
import csv
import time
import datetime as dt
from pathlib import Path
from glob import glob
from typing import List, Optional, Tuple

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
CSV_FILE = "reports/project.csv"

# Pinned URLs and role for Project Manager scenario
REPORT_URL = "https://officemanager.dodopizza.ru/OfficeManager/Debiting/PrepareExcelReport"
SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"
BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"
ROLE_ID = "8"

# Optional slow delay can still be overridden via env
SLOW_DELAY = float(os.environ.get("SLOW_DELAY", "0"))


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

    return webdriver.Chrome(service=service, options=options)


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
        self.driver.get(self.select_department_url)
        self.ensure_role_selected()
        if "/SelectDepartment" not in self.driver.current_url:
            try:
                self.driver.get(self.select_department_url)
            except Exception:
                pass

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

    # ---------- CSV ----------
    def reset_csv(self) -> None:
        with open(self.csv_file, "w", encoding="utf-8-sig", newline="") as f:
            f.write("\ufeff")

    def append_csv_row(self, row: List[str]) -> None:
        with open(self.csv_file, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter=";").writerow(row)

    # ---------- Main flow ----------
    def run(self) -> int:
        self.open_select_department()
        cities = self.get_cities()
        print(f"[CITIES] Найдено: {len(cities)} — {', '.join([c[0] for c in cities])}")

        dates = self.compute_dates()
        print(
            f"[DATES] Диапазон: {dates[0]:%d.%m.%Y} — {dates[-1]:%d.%m.%Y} (всего {len(dates)})"
            if dates
            else "[DATES] Сегодня 1-е — диапазон пуст"
        )

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
                self.append_csv_row([f"ГОРОД: {city_name}", ""])  # header

                for didx, dept in enumerate(departments, start=1):
                    print("\n" + "=" * 80)
                    print(f"[DEPT] ({didx}/{len(departments)}) {dept}")
                    self.choose_department(dept)
                    self.append_csv_row([f"ОТДЕЛ: {dept}", ""])  # section

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
                        self.append_csv_row([d.strftime("%d.%m.%Y"), val])
                        print(f"[CSV] {d:%d.%m.%Y}: {val}")

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
    user_data_dir = Path(os.environ.get("USER_DATA_DIR", "/profile"))
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
