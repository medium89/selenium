# -*- coding: utf-8 -*-
# Скрипт открывает страницу резюме, извлекает пары "метка: значение"
# из двухколоночных блоков и сохраняет их в CSV.

# Импорты стандартной библиотеки
import csv
import re
import time

# Импорты Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

START_URL = "https://ivanovdo.ru/resume"  # Конечная страница для парсинга (при необходимости замените на свою)

def norm(text: str) -> str:
    """Нормализация текста.

    Блок: приведение текста к удобному виду для записи/сравнения:
    - заменяем неразрывные пробелы на обычные;
    - сжимаем все последовательности пробелов до одного;
    - обрезаем пробелы по краям.
    """
    text = text.replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()

def main():
    # Блок: инициализация браузера и ожиданий
    d = webdriver.Chrome()  # Создаём драйвер Chrome (можно заменить на другой при необходимости)
    d.set_window_size(1440, 900)  # Фиксируем размер окна, чтобы элементы были в зоне видимости
    wait = WebDriverWait(d, 20)  # Явные ожидания до 20 секунд

    try:
        # Блок: переход на целевую страницу
        d.get(START_URL)

        # Дожидаемся полной загрузки документа и проверяем выполнение JS (alert + DOM-маркер)
        try:
            WebDriverWait(d, 20).until(lambda drv: drv.execute_script("return document.readyState") == "complete")
            # Alert для визуальной проверки выполнения JS
            d.execute_script("alert('JS работает: страница полностью загружена');")
            try:
                WebDriverWait(d, 5).until(EC.alert_is_present())
                print("Alert показан. Закрою через 3 секунды…")
                time.sleep(3)
                d.switch_to.alert.accept()
            except Exception:
                pass
        except Exception as e:
            print(f"Не удалось показать alert: {e}")

        # DOM-маркер в правом верхнем углу
        try:
            inserted = d.execute_script(
                """
                (function(){
                  try{
                    var id='__selenium_js_marker__';
                    var el=document.getElementById(id);
                    if(!el){
                      el=document.createElement('div');
                      el.id=id; el.textContent='SELENIUM JS OK';
                      var s=el.style; s.position='fixed'; s.top='10px'; s.right='10px'; s.zIndex='2147483647';
                      s.background='#28a745'; s.color='#fff'; s.padding='8px 12px'; s.fontFamily='Arial, sans-serif'; s.fontSize='14px';
                      s.borderRadius='4px'; s.boxShadow='0 2px 6px rgba(0,0,0,0.3)';
                      document.body.appendChild(el);
                    }
                    return !!el;
                  }catch(e){ return false; }
                })();
                """
            )
            print(f"DOM-маркер добавлен: {inserted}")
        except Exception as e:
            print(f"Не удалось добавить DOM-маркер: {e}")

        # Блок: ожидание появления структурных блоков строк (row)
        wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//div[contains(@class,'row')]")
        ))

        # Блок: выбор только тех строк, где есть левая метка и правое значение
        rows = d.find_elements(
            By.XPATH,
            "//div[contains(@class,'row')][div[contains(@class,'col-md-3') and contains(@class,'control-label')]"
            " and div[contains(@class,'col-md-9')]]"
        )

        # Блок: подготовка CSV для вывода результатов
        out_path = "rows.csv"
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)

            # Блок: обход всех подходящих строк и извлечение пары "метка: значение"
            for row in rows:
                # Левый блок (метка)
                left = row.find_element(
                    By.XPATH, "./div[contains(@class,'col-md-3') and contains(@class,'control-label')]"
                ).text
                # Правый блок (значение)
                right = row.find_element(
                    By.XPATH, "./div[contains(@class,'col-md-9')]"
                ).text

                # Блок: нормализация значений и лёгкая очистка
                left = norm(left).rstrip(":")   # убираем двоеточие на конце, если есть
                right = norm(right)

                # Блок: фильтрация пустых строк и запись результата
                if left and right:              # пропускаем пустые/служебные строки
                    w.writerow([f"{left}: {right}"])

        # Блок: сообщение об успешном завершении
        print(f"Готово! Сохранено в {out_path}")

    finally:
        # Блок: корректное завершение сессии браузера даже при ошибках
        d.quit()

if __name__ == "__main__":
    # Точка входа: запускаем основную функцию при прямом вызове скрипта
    main()
