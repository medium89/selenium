import os
import sys
import subprocess
from pathlib import Path


SCRIPTS = {
    "1": ("OfficeManager — MaterialConsumption (отделы/даты → reports/office.csv)", "OfficeManager.py"),
    "2": ("ProjectManager — Debiting/PrepareExcelReport (итоги → reports/project.csv)", "ProjectManager.py"),
    "3": ("RevisiaMetrics — LossesAndExcees (Статистика/ревизии → reports/revisia.csv)", "RevisiaMetrics.py"),
    "4": ("OpenSelectRole — просто открыть выбор роли", "OpenSelectRole.py"),
    "5": ("AnalyticsSpy — сохранить HTML отчёта в spizdil.html", "AnalyticsSpy.py"),
    "6": ("AnalyticsLostRevenue — собрать ‘Доля упущенной выручки’ по городам", "AnalyticsLostRevenue.py"),
    "7": (
        "AnalyticsKeyMetrics — собрать 'Доля упущенной выручки по стопам пиццерии' (Итого)",
        "AnalyticsKeyMetrics.py",
    ),
}


def main() -> int:
    print("Выберите скрипт для запуска:")
    keys_sorted = sorted(SCRIPTS.keys(), key=int)
    for k in keys_sorted:
        title, _ = SCRIPTS[k]
        print(f" {k}. {title}")
    print(" q. Выход")

    max_choice = max((int(k) for k in keys_sorted), default=0)
    if max_choice:
        choice_prompt = f"Ваш выбор (1-{max_choice}/q): "
        error_hint = f"Введите 1-{max_choice} или q для выхода."
    else:
        choice_prompt = "Ваш выбор (q): "
        error_hint = "Введите q для выхода."

    while True:
        try:
            choice = input(choice_prompt).strip()
        except EOFError:
            choice = "q"
        if choice.lower() in {"q", "quit", "exit"}:
            return 0
        if choice in SCRIPTS:
            title, rel_path = SCRIPTS[choice]
            script_path = Path(__file__).parent / rel_path
            if not script_path.exists():
                print(f"[error] Не найден файл: {script_path}")
                return 2
            print(f"[run] {title}")
            env = os.environ.copy()
            cmd = [sys.executable, str(script_path)]
            try:
                return subprocess.call(cmd, env=env)
            except KeyboardInterrupt:
                return 130
            except Exception as e:
                print(f"[error] Не удалось запустить {rel_path}: {e}")
                return 3
        else:
            print(error_hint)


if __name__ == "__main__":
    raise SystemExit(main())
