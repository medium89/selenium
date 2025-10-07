Selenium-скрипты и запуск через .venv

Коротко: в `app/` лежат скрипты для автоматизации Chrome через Selenium. Они собирают отчёты и сохраняют CSV в `./reports`. По умолчанию скрипты запускаются в headless и используют профиль `./profile`. Есть утилита для открытия страницы выбора роли с окном и меню запуска скриптов.

Состав app/ и результаты
- `app/OfficeManager.py`: MaterialConsumption по городам/отделам/датам → `reports/office.csv`.
- `app/ProjectManager.py`: Debiting/PrepareExcelReport (итоги по дням) → `reports/project.csv`.
- `app/RevisiaMetrics.py`: LossesAndExcees/«Статистика» с перебором ревизий → `reports/revisia.csv`.
- `app/OpenSelectRole.py`: открывает страницу SelectRole и печатает доступные роли (GUI по умолчанию, окно не закрывает).
- `app/init.py`: консольное меню выбора и запуска одного из скриптов.
- `app/AnalyticsSpy.py`: сохраняет HTML страницы отчёта «Бизнес‑обзор/Аналитика» в `spizdil.html` (для отладки селекторов).
- `app/AnalyticsLostRevenue.py`: собирает метрику «Доля упущенной выручки» по всем городам → `reports/lost_revenue.csv`.

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
- `python app/OfficeManager.py`
- `python app/ProjectManager.py`
- `python app/RevisiaMetrics.py`
- `python app/OpenSelectRole.py` — открыть SelectRole (GUI, окно не закрывается)
- `python app/init.py` — меню выбора скрипта
— Дополнительно:
- `python app/AnalyticsSpy.py` — сохранить HTML отчёта в `spizdil.html`
- `python app/AnalyticsLostRevenue.py` — собрать «Доля упущенной выручки» по всем городам

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
- С окном: `HEADLESS=0 python app/OfficeManager.py` (или откройте `app/OpenSelectRole.py`).
- Без окна: по умолчанию headless включён.
- Если сайт отдаёт «Forbidden» в headless, используйте GUI или подключение к уже открытому Chrome: `chromium --user-data-dir="$PWD/profile" --remote-debugging-port=9222 about:blank`, затем запустите скрипт — он прицепится к открытому окну.

Где искать результаты
- `reports/office.csv` — MaterialConsumption.
- `reports/project.csv` — Debiting/PrepareExcelReport.
- `reports/revisia.csv` — LossesAndExcees/«Статистика».
— Дополнительно: `reports/lost_revenue.csv` — «Доля упущенной выручки».
CSV — UTF‑8 с BOM, разделитель `;`.

Analytics: метрика «Доля упущенной выручки»
- Сохранить HTML дашборда для отладки: `python app/AnalyticsSpy.py` (файл `spizdil.html` в корне).
- Собрать метрику по всем городам: `python app/AnalyticsLostRevenue.py` (CSV: `reports/lost_revenue.csv`).
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
  - `python app\OfficeManager.py`
- Авторизуйтесь в открывшемся окне Chrome. Закройте его и снова запустите Docker‑команду.

**Значения по умолчанию (ProjectManager)**
- В `app/ProjectManager.py` зашиты ключевые константы:
  - `REPORT_URL = "https://officemanager.dodopizza.ru/OfficeManager/Debiting/PrepareExcelReport"`
  - `SELECT_DEPARTMENT_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"`
  - `BACK_TO_SELECT_ROLE_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/BackToSelectRole"`
  - `ROLE_ID = "8"`
  - `CSV_FILE = "reports/project.csv"`
- Для изменения поведения отредактируйте эти константы и пересоберите образ.

**Диагностика**
- Профиль занят: «user data directory is already in use»
  - Закройте Chrome; удалите локи в `./profile` (команда выше); повторите.
- Предупреждение Compose «version is obsolete» — можно игнорировать.
- В CSV пустые суммы:
  - Убедитесь, что вы авторизованы и на странице видны итоги.
  - Если разметка отчёта иная, поправьте селекторы в `read_total_value()` внутри `app/ProjectManager.py`.
