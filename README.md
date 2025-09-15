**Обзор**
- Скрипты в `app/` автоматизируют Chrome через Selenium.
- `app/OfficeManager.py`: открывает страницу и сохраняет таблицу в CSV.
- `app/ProjectManager.py`: строит отчёты по городам/отделам и сохраняет CSV.
- Docker‑образ содержит Chromium и Chromedriver, запускается в headless и использует локальный профиль из `./profile`.

**Требования**
- Docker Desktop (Windows/macOS/Linux).
- Локальный профиль Chrome в `./profile` с действующей авторизацией для `officemanager.dodopizza.ru` (см. ниже «Разовый вход через GUI»).

**Структура**
- `app/OfficeManager.py`: сохранение таблицы со страницы в `./reports/office.csv` (если таблицы нет — пишется заголовок страницы).
- `app/ProjectManager.py`: генерация отчёта в `./reports/project.csv`.
- `docker/Dockerfile`: образ с Chromium + Chromedriver.
- `docker/docker-compose.yml`: сервис `selenium-app` и тома:
  - `./profile:/profile` (профиль браузера)
  - `./reports:/app/reports` (выгрузки CSV)
- `docker/docker_init.id`: быстрые команды (PowerShell).

**Сборка образа**
- PowerShell (из корня репозитория):
  - `docker compose -f docker/docker-compose.yml build`
  - Чистая пересборка: `docker compose -f docker/docker-compose.yml build --no-cache`

**Запуск OfficeManager (CSV)**
- Открыть целевой URL и сохранить таблицу в `./reports/office.csv`:
- PowerShell:
  - `# опционально; по умолчанию https://www.google.com`
  - `$env:START_URL = "https://officemanager.dodopizza.ru/Infrastructure/Authenticate/SelectDepartment"`
  - `docker compose -f docker/docker-compose.yml run --rm selenium-app`

**Запуск ProjectManager (CSV)**
- Сохраняет в `./reports/project.csv`.
- PowerShell:
  - `docker compose -f docker/docker-compose.yml run --rm selenium-app python app/ProjectManager.py`
- Проверить первые строки:
  - `Get-Content .\reports\project.csv -TotalCount 20`

**Профиль и авторизация**
- Контейнер монтирует `./profile` в `/profile` и запускает Chrome с `--user-data-dir=/profile`.
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
