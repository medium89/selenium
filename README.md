Selenium-скрипты и запуск через .venv

Коротко: в `app/` лежат скрипты для автоматизации Chrome через Selenium. Они собирают отчёты и сохраняют CSV в `./reports`. По умолчанию скрипты запускаются в headless и используют профиль `./profile`. Есть утилита для открытия страницы выбора роли с окном и меню запуска скриптов.

Состав app/ и результаты
- `app/MaterialConsumptionReport.py`: MaterialConsumption по городам/отделам/датам → `reports/MaterialConsumptionReport.csv`.
- `app/MaterialWriteOffReport.py`: Debiting/PrepareExcelReport (итоги по дням) → `reports/MaterialWriteOffReport.csv`.
- `app/LossesAndExcessReport.py`: LossesAndExcees/«Статистика» с перебором ревизий → `reports/LossesAndExcessReport.csv`.
- `app/AnalyticsLostRevenue.py`: «Доля упущенной выручки» по городам/пиццериям → `reports/AnalyticsLostRevenue.csv`.
- `app/init.py`: консольное меню с пунктами «Загрузить отчёт …» и подменю «Технические скрипты».
- `app/OpenSelectRole.py`: открывает SelectRole и печатает доступные роли (GUI по умолчанию).
- `app/AnalyticsSpy.py`: сохраняет HTML дашборда «Analytics» в `parsing.html` для отладки селекторов.

Подготовка .venv
- Windows (PowerShell):
  - `py -3 -m venv .venv`
  - `.\.venv\Scripts\Activate.ps1`
  - `pip install -r config\requirements.txt`
- Linux/macOS (bash/zsh):
  - `python3 -m venv .venv`
  - `source .venv/bin/activate`
  - `pip install -r config/requirements.txt`

Запуск (дефолты зашиты)
- `python app/MaterialConsumptionReport.py`
- `python app/MaterialWriteOffReport.py`
- `python app/LossesAndExcessReport.py`
- `python app/AnalyticsLostRevenue.py`
- `python app/init.py` — меню; отчёты перечислены как «Загрузить отчёт: …», технические утилиты лежат в подменю.
— Дополнительно:
- `python app/OpenSelectRole.py` — открыть SelectRole (GUI, окно не закрывается)
- `python app/AnalyticsSpy.py` — сохранить HTML отчёта в `parsing.html`

Дефолты, зашитые в скриптах
- Профиль: `USER_DATA_DIR=$PWD/profile` (создаётся, если нет).
- Режим: `HEADLESS=1` (headless) для рабочих скриптов; у `OpenSelectRole.py` — `HEADLESS=0` по умолчанию.
- Stealth: `STEALTH=1` (подмена UA, lang, таймзона, JS‑патчи, отключение AutomationControlled).

Расширенный ввод дат
- Пустой ввод — вчерашняя дата.
- `04` — 4 число текущего месяца и года.
- `04102025` — 04.10.2025.
- Также поддерживаются `ДД.ММ.ГГГГ`, `ДД‑ММ‑ГГГГ`, `ГГГГ‑ММ‑ДД`.

Chrome с GUI и без
- С окном: `HEADLESS=0 python app/MaterialConsumptionReport.py` (или откройте `app/OpenSelectRole.py`).
- Без окна: по умолчанию headless включён.
- Если сайт отдаёт «Forbidden» в headless, используйте GUI или подключение к уже открытому Chrome: `chromium --user-data-dir="$PWD/profile" --remote-debugging-port=9222 about:blank`, затем запустите скрипт — он прицепится к открытому окну.

Где искать результаты
- `reports/MaterialConsumptionReport.csv` — MaterialConsumption.
- `reports/MaterialWriteOffReport.csv` — Debiting/PrepareExcelReport.
- `reports/LossesAndExcessReport.csv` — LossesAndExcees/«Статистика».
- `reports/AnalyticsLostRevenue.csv` — «Доля упущенной выручки».
- `reports/bugrepot.txt` — журнал ответов Supabase (успехи/ошибки по каждому чанку).
CSV — UTF‑8 с BOM, разделитель `;`.

Supabase и логирование
- Все отчёты (MaterialConsumption, MaterialWriteOff, LossesAndExcess, AnalyticsLostRevenue) отправляют данные в Supabase, если заданы `SUPABASE_URL`/`SUPABASE_KEY` (или `SUPABASE_SERVICE_KEY`). Таблица по умолчанию берётся из `SUPABASE_TABLE`.
- Конфигурацию можно положить в `config/api` (JSON или `key=value`), ключ — в `config/api.key`, либо передать переменными окружения.
- При указании `SUPABASE_ON_CONFLICT` скрипты делают `INSERT ... on_conflict=...`. Если Supabase возвращает ошибку «no unique or exclusion constraint…», скрипт автоматически повторит запрос без on_conflict и запишет предупреждение в bugreport.
- Полный ответ curl (stdout/stderr, размер чанка) попадает в `reports/bugrepot.txt`. Ошибки Supabase больше не теряются — проверяйте файл после выгрузки.
- Для `AnalyticsLostRevenue.py` значения `SUPABASE_TABLE` по умолчанию `analytics_lost_revenue`, `SUPABASE_ON_CONFLICT=city,department,report_date`. Пример SQL для таблицы:
  ```sql
  BEGIN;
  DROP TABLE IF EXISTS public.analytics_lost_revenue;
  CREATE TABLE public.analytics_lost_revenue (
      id              bigserial PRIMARY KEY,
      city            text        NOT NULL,
      department      text        NOT NULL,
      report_date     date        NOT NULL,
      lost_share      numeric,
      created_at      timestamptz NOT NULL DEFAULT timezone('utc', now()),
      updated_at      timestamptz NOT NULL DEFAULT timezone('utc', now())
  );
  CREATE INDEX analytics_lost_revenue_report_date_idx ON public.analytics_lost_revenue (report_date);
  CREATE INDEX analytics_lost_revenue_city_idx        ON public.analytics_lost_revenue (city);
  CREATE INDEX analytics_lost_revenue_department_idx ON public.analytics_lost_revenue (department);
  COMMIT;
  ```
- Для LossesAndExcess не забудьте уникальный индекс под `SUPABASE_ON_CONFLICT` (по умолчанию `city,department,dt,revisions,metric_name`).

Analytics: метрика «Доля упущенной выручки»
- Сохранить HTML дашборда для отладки: `python app/AnalyticsSpy.py` (файл `parsing.html` в корне).
- Собрать метрику по всем городам: `python app/AnalyticsLostRevenue.py` (CSV: `reports/AnalyticsLostRevenue.csv`).
- Если фильтры не применяются в headless, откройте Chrome с окном и портом: `chromium --user-data-dir="$PWD/profile" --remote-debugging-port=9225 about:blank`, затем запустите скрипт — он подключится к открытому окну.

Быстрая диагностика
- «user data directory is already in use» — закройте Chrome и удалите lock‑файлы в `profile/` (`Singleton*`, `DevToolsActivePort`).
- «Forbidden» в headless — включите окно (`HEADLESS=0`) или подключайтесь к открытому Chrome (см. выше).
- Нет данных — проверьте авторизацию профиля и актуальность роли/селекторов.
- Если редиректит на `auth.dodois.io`, профиль не авторизован — выполните разовый вход через GUI локально, затем повторите запуск в Docker.

**Разовый вход через GUI (локально, не в Docker)**
- Закройте все окна Chrome и очистите локи профиля:
  - `Get-ChildItem -Path .\profile -Recurse -Force -Include "Singleton*", "DevToolsActivePort" | Remove-Item -Force -ErrorAction SilentlyContinue`
- Создайте виртуальное окружение и установите зависимости:
  - `py -3.11 -m venv .venv`
  - `.\.venv\Scripts\Activate.ps1`
  - `pip install -r config\requirements.txt`
- Запуск с GUI для входа:
  - `$env:HEADLESS = "0"`
  - `$env:USER_DATA_DIR = "$PWD\profile"`
  - `$env:LOGIN_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"`
  - `python app\MaterialConsumptionReport.py`
- Авторизуйтесь в открывшемся окне Chrome. Закройте его и снова запустите Docker‑команду.

**Значения по умолчанию (Списания сырья)**
- В `app/MaterialWriteOffReport.py` зашиты ключевые константы:
  - `REPORT_URL = "https://officemanager.dodopizza.ru/OfficeManager/Debiting/PrepareExcelReport"`
  - `SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"`
  - `BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"`
  - `ROLE_ID = "8"`
  - `CSV_FILE = "reports/MaterialWriteOffReport.csv"`
- Для изменения поведения отредактируйте эти константы и пересоберите образ.

**Диагностика**
- Профиль занят: «user data directory is already in use»
  - Закройте Chrome; удалите локи в `./profile` (команда выше); повторите.
- Предупреждение Compose «version is obsolete» — можно игнорировать.
- В CSV пустые суммы:
  - Убедитесь, что вы авторизованы и на странице видны итоги.
  - Если разметка отчёта иная, поправьте селекторы в `read_total_value()` внутри `app/MaterialWriteOffReport.py`.
