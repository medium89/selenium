from __future__ import annotations

import datetime as dt
import os
import time
from typing import Optional

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

from AnalyticsLostRevenue import (  # reuse helper tools and auth flow
    apply_filters,
    build_driver,
    _clear_tokens_within_select,
    _close_dropdown,
    _find_select_by_label,
    _get_visible_dropdown,
    _open_select_dropdown,
    _dropdown_wait_options,
    ensure_on_analytics_or_retry,
    get_cities,
    get_pizzerias_for_city,
    goto_analytics,
    open_filter_panel,
    select_city,
    set_country_russia_only,
    set_pizzerias_for_city,
    wait_filters_loaded,
    write_csv_row,
)

ROW_LABEL = os.environ.get("KEY_METRIC_ROW", "Итого")
COLUMN_LABEL = os.environ.get("KEY_METRIC_COLUMN", "Доля упущенной выручки по стопам пиццерии")
COLUMN_ARIA_NEEDLE = os.environ.get("KEY_METRIC_COLUMN_ARIA", "lost_revenue_pizz").strip().lower()
TABLE_SELECTOR = os.environ.get(
    "KEY_METRIC_TABLE_SELECTOR",
    "table.table.table-striped.table-condensed, table.table-striped.table-condensed",
)
CSV_FILE = os.environ.get("KEY_METRIC_CSV", "reports/key_metric_totals.csv")


def clear_pizzeria_filter(driver: webdriver.Chrome) -> None:
    sel = _find_select_by_label(driver, ["пиццерия", "store", "pizzeria", "пиццер"])
    if not sel:
        return
    # Remove visible tokens / use clear button
    _clear_tokens_within_select(driver, sel)
    try:
        clears = sel.find_elements(By.CSS_SELECTOR, ".ant-select-clear")
    except Exception:
        clears = []
    for btn in clears:
        try:
            btn.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click()", btn)
            except Exception:
                pass

    # Open dropdown and toggle select-all / selected options off (covers virtualization cases)
    try:
        _open_select_dropdown(driver, sel)
        _dropdown_wait_options(driver, timeout=6.0)
        dd = _get_visible_dropdown(driver)
    except Exception:
        dd = None

    toggled = False
    if dd:
        try:
            select_all_candidates = dd.find_elements(
                By.CSS_SELECTOR,
                ".select-all, [id*='select-all'], [data-test='native-filters-select-all']",
            )
        except Exception:
            select_all_candidates = []
        for opt in select_all_candidates:
            try:
                aria_selected = (opt.get_attribute("aria-selected") or "").lower()
            except Exception:
                aria_selected = ""
            should_toggle = True
            if aria_selected in {"false", "0"}:
                # Some dashboards mark select-all inactive via class
                try:
                    classes = opt.get_attribute("class") or ""
                    if "ant-select-item-option-selected" not in classes and "active" not in classes:
                        should_toggle = False
                except Exception:
                    pass
            if should_toggle:
                try:
                    opt.click()
                    toggled = True
                except Exception:
                    try:
                        driver.execute_script("arguments[0].click()", opt)
                        toggled = True
                    except Exception:
                        pass
    try:
        _close_dropdown(driver)
    except Exception:
        pass

    if toggled:
        # Give React time to sync and remove tags
        time.sleep(0.4)
        _clear_tokens_within_select(driver, sel)

    # Final safeguard: reset search input state
    try:
        driver.execute_script(
            """
            const root = arguments[0];
            if (!root) return;
            const inputs = root.querySelectorAll('input.ant-select-selection-search-input');
            inputs.forEach(input => {
                input.value = '';
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
            });
            """,
            sel,
        )
    except Exception:
        pass


def read_key_metric_total(
    driver: webdriver.Chrome,
    row_label: str = ROW_LABEL,
    column_label: str = COLUMN_LABEL,
) -> str:
    script = """
        const normalize = (txt) => {
            if (txt == null) {
                return "";
            }
            return String(txt).replace(/\\u00a0/g, " ").replace(/\\s+/g, " ").trim().toLowerCase();
        };
        const highlightCell = (cell) => {
            if (!cell) return;
            try {
                cell.style.outline = "2px solid #6d28d9";
                cell.style.backgroundColor = "#ede9fe";
                cell.scrollIntoView({ block: "center", inline: "nearest" });
            } catch (e) {
                // ignore
            }
        };
        const rowNeedle = normalize(arguments[0]);
        const colNeedle = normalize(arguments[1]);
        const colAriaNeedle = String(arguments[2] || "").toLowerCase();
        const tableSelector = arguments[3] || "";
        const tables = Array.from(document.querySelectorAll("table")).filter((tbl) => {
            if (!tableSelector) return true;
            try {
                return tbl.matches(tableSelector);
            } catch (e) {
                return true;
            }
        });

        const extractCellValue = (cell) => {
            if (!cell) return "";
            const raw = (cell.textContent || "").replace(/\\u00a0/g, " ").trim();
            return raw;
        };

        const resolveCell = (cells) => {
            let fallback = null;
            for (const cell of cells) {
                const ariaIds = (cell.getAttribute("aria-labelledby") || cell.getAttribute("headers") || "")
                    .split(/\\s+/)
                    .filter(Boolean);
                let headerMatched = false;
                for (const ariaId of ariaIds) {
                    const headerEl = document.getElementById(ariaId);
                    if (!headerEl) continue;
                    const headerText = normalize(headerEl.textContent);
                    if (colNeedle && headerText.includes(colNeedle)) {
                        highlightCell(cell);
                        return cell;
                    }
                    if (!fallback && colAriaNeedle && ariaId.toLowerCase().includes(colAriaNeedle)) {
                        fallback = cell;
                    }
                }
                const dataLabel = normalize(cell.getAttribute("data-label"));
                if (colNeedle && dataLabel.includes(colNeedle)) {
                    highlightCell(cell);
                    return cell;
                }
                const titleAttr = normalize(cell.getAttribute("title"));
                if (colNeedle && titleAttr.includes(colNeedle)) {
                    highlightCell(cell);
                    return cell;
                }
                if (!fallback && colAriaNeedle) {
                    const ariaJoined = ariaIds.join(" ").toLowerCase();
                    if (ariaJoined.includes(colAriaNeedle)) {
                        fallback = cell;
                    }
                }
            }
            if (fallback) {
                highlightCell(fallback);
            }
            return fallback;
        };

        const findRowMatch = (rows, expectedRow) => {
            let firstRow = null;
            for (const row of rows) {
                const cells = Array.from(row.querySelectorAll("th,td"));
                if (!cells.length) {
                    continue;
                }
                if (!firstRow) {
                    firstRow = { row, cells };
                }
                if (!expectedRow) {
                    return { row, cells };
                }
                const firstCell = cells[0];
                const rowHeaderText = normalize(firstCell.textContent);
                const fullRowText = normalize(row.textContent);
                if (
                    (rowHeaderText && rowHeaderText.includes(expectedRow)) ||
                    (fullRowText && fullRowText.includes(expectedRow))
                ) {
                    return { row, cells };
                }
            }
            return firstRow;
        };

        for (const table of tables) {
            const rows = Array.from(table.querySelectorAll("tbody tr, tr"));
            if (!rows.length) {
                continue;
            }
            const targetRow = findRowMatch(rows, rowNeedle);
            if (!targetRow) {
                continue;
            }
            const dataCells = targetRow.cells.filter(
                (cell) => cell.tagName && cell.tagName.toLowerCase() === "td"
            );
            const candidates = dataCells.length ? dataCells : targetRow.cells;
            const cell = resolveCell(candidates);
            if (cell) {
                const raw = extractCellValue(cell);
                if (raw) {
                    return raw;
                }
            }
        }
        return null;
    """
    try:
        value = driver.execute_script(script, row_label, column_label, COLUMN_ARIA_NEEDLE, TABLE_SELECTOR)
    except Exception:
        return ""
    if not value:
        return ""
    return str(value).strip()


def wait_key_metric_update(
    driver: webdriver.Chrome,
    row_label: str,
    column_label: str,
    previous: Optional[str] = None,
    timeout: float = 35.0,
    poll: float = 0.6,
) -> str:
    min_wait = time.time() + 2.0
    deadline = time.time() + max(timeout, 1.0)
    last_val: Optional[str] = None
    stable_reads = 0
    prev_clean = (previous or "").strip()

    while time.time() < deadline:
        current = (read_key_metric_total(driver, row_label, column_label) or "").strip()
        if not current:
            last_val = None
            stable_reads = 0
            time.sleep(poll)
            continue

        if current == last_val:
            stable_reads += 1
        else:
            last_val = current
            stable_reads = 1

        if stable_reads >= 2:
            if prev_clean and current == prev_clean:
                if time.time() >= min_wait:
                    return current
            else:
                if time.time() >= min_wait:
                    return current
        time.sleep(poll)

    return last_val or prev_clean or ""


def main() -> int:
    driver = build_driver()
    wait = WebDriverWait(driver, 25)
    today = dt.date.today().strftime("%d.%m.%Y")

    try:
        cities = get_cities(driver, wait)
        print(f"[CITIES] Найдено: {len(cities)} — {', '.join([c[0] for c in cities])}")
        for idx, (city_name, city_uuid) in enumerate(cities, start=1):
            print("\n" + "#" * 80)
            print(f"[CITY] ({idx}/{len(cities)}) {city_name}")
            try:
                select_city(driver, wait, city_uuid)
                goto_analytics(driver, wait)
                try:
                    WebDriverWait(driver, 30).until(
                        lambda d: d.execute_script("return document.readyState") == "complete"
                    )
                except Exception:
                    pass
                ensure_on_analytics_or_retry(driver, wait, retries=2)

                prev_metric_city: Optional[str] = None
                try:
                    open_filter_panel(driver)
                    wait_filters_loaded(driver, wait, timeout=30)
                    set_country_russia_only(driver)

                    city_key = (city_name or "").strip().lower()
                    if city_key == "ставрополь":
                        pizzerias = get_pizzerias_for_city(driver, city_name) or []
                        if not pizzerias:
                            pizzerias = [city_name]
                        for branch_idx, branch_name in enumerate(pizzerias, start=1):
                            if branch_idx > 1:
                                open_filter_panel(driver)
                                wait_filters_loaded(driver, wait, timeout=30)
                                set_country_russia_only(driver)
                            try:
                                clear_pizzeria_filter(driver)
                                selected = set_pizzerias_for_city(
                                    driver, city_name, exact_pizzerias=[branch_name]
                                )
                                if not selected:
                                    raise RuntimeError("пиццерия не была выбрана в фильтре")
                                prev_metric_branch = read_key_metric_total(driver)
                                apply_filters(driver)
                                print(f"[FILTER] Применены фильтры для {branch_name}")
                                branch_val = wait_key_metric_update(
                                    driver, ROW_LABEL, COLUMN_LABEL, previous=prev_metric_branch
                                )
                                write_csv_row(CSV_FILE, [branch_name, today, branch_val or "(нет данных)"])
                                print(f"[CSV] {branch_name}: {branch_val}")
                            except Exception as branch_err:
                                print(f"[WARN] Ошибка в пиццерии {branch_name}: {branch_err}")
                                write_csv_row(
                                    CSV_FILE, [branch_name, today, f"ОШИБКА: {branch_err}"]
                                )
                        continue

                    clear_pizzeria_filter(driver)
                    set_pizzerias_for_city(driver, city_name)
                    prev_metric_city = read_key_metric_total(driver)
                    apply_filters(driver)
                    print(f"[FILTER] Применены фильтры для {city_name}")
                except Exception as e:
                    print(f"[filters] Не удалось применить фильтры: {e}")

                val = wait_key_metric_update(
                    driver,
                    ROW_LABEL,
                    COLUMN_LABEL,
                    previous=prev_metric_city,
                )
                write_csv_row(CSV_FILE, [city_name, today, val or "(нет данных)"])
                print(f"[CSV] {city_name}: {val}")
            except Exception as city_err:
                print(f"[WARN] Ошибка в городе {city_name}: {city_err}")
                write_csv_row(CSV_FILE, [city_name, today, f"ОШИБКА: {city_err}"])
    finally:
        try:
            driver.quit()
        except Exception:
            pass
    print(f"[DONE] Готово! Файл {CSV_FILE} сохранён.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
