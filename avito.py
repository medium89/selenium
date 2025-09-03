# avito_reviews_scrape.py
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

URL = "https://www.avito.ru/barnaul/predlozheniya_uslug/vygul_sobak_zoonyanya_perederzhka_uhod_za_pitomtsem_7425878726?context=H4sIAAAAAAAA_wEfAOD_YToxOntzOjEzOiJsb2NhbFByaW9yaXR5IjtiOjA7fQseF2QfAAAA#open-reviews-list"

def main():
    # 1) Настраиваем Chrome
    options = webdriver.ChromeOptions()
    # Если нужно без окна браузера — раскомментируй:
    # options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # Немного маскировки, помогает на некоторых сайтах
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 15)

    try:
        # 2) Открываем страницу
        driver.get(URL)

        # 3) Закрываем возможные попапы (cookies/регион) — необязательно, но полезно
        for selector in [
            "button[aria-label='Принять']",         # cookies (часто)
            "button[data-marker='popuptablet/accept']",
            "button[data-marker='regionConfirmation/accept']",
            "button[aria-label='Закрыть']"
        ]:
            try:
                btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                btn.click()
                time.sleep(0.5)
            except TimeoutException:
                pass

        # 4) Переходим к секции с отзывами (якорь уже в URL, но прокрутимся)
        driver.execute_script("location.hash = '#open-reviews-list';")
        time.sleep(0.5)

        # 5) Ждём появления хотя бы одного нужного контейнера отзывов
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.aXMKx")))

        # 6) Прокрутка для подгрузки (если отзывы лениво грузятся)
        # Сделаем несколько шагов вниз, собирая новые элементы
        last_height = driver.execute_script("return document.body.scrollHeight")
        for _ in range(6):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1.2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # 7) Собираем тексты отзывов из контейнеров <div class="aXMKx">
        review_divs = driver.find_elements(By.CSS_SELECTOR, "div.aXMKx")
        reviews = []
        for div in review_divs:
            txt = div.text.strip()
            if txt:
                reviews.append(txt)

        # 8) Если на странице есть кнопка "Показать ещё" для отзывов — попытаемся нажать пару раз
        # (на некоторых страницах встречается)
        for _ in range(3):
            try:
                more_btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(),'Показать ещё') or contains(text(),'Ещё')]]"))
                )
                driver.execute_script("arguments[0].click();", more_btn)
                time.sleep(1.2)
                # добираем новые блоки
                new_divs = driver.find_elements(By.CSS_SELECTOR, "div.aXMKx")
                for div in new_divs:
                    txt = div.text.strip()
                    if txt and txt not in reviews:
                        reviews.append(txt)
            except TimeoutException:
                break

        # 9) Сохраняем в файл
        with open("avito_reviews.txt", "w", encoding="utf-8") as f:
            for r in reviews:
                f.write(r.replace("\r", "").strip() + "\n\n---\n\n")

        print(f"Собрано отзывов: {len(reviews)}")
        for i, r in enumerate(reviews, 1):
            print(f"[{i}] {r}\n")

        if not reviews:
            print("Предупреждение: отзывы не найдены. Проверь, не изменился ли класс контейнера или нужна ли авторизация/скролл.")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
