import os
import sys
import subprocess
from pathlib import Path
from typing import Dict, Optional, Tuple, Union


SCRIPTS = {
    "1": ("Загрузить отчёт: Расход сырья (MaterialConsumptionReport.py)", "MaterialConsumptionReport.py"),
    "2": ("Загрузить отчёт: Списания сырья (MaterialWriteOffReport.py)", "MaterialWriteOffReport.py"),
    "3": ("Загрузить отчёт: Потери и избыток (LossesAndExcessReport.py)", "LossesAndExcessReport.py"),
    "4": ("Загрузить отчёт: Доля упущенной выручки (AnalyticsLostRevenue.py)", "AnalyticsLostRevenue.py"),
    "5": (
        "Технические скрипты",
        {
            "1": ("Технический скрипт: открыть выбор роли (OpenSelectRole.py)", "OpenSelectRole.py"),
            "2": ("Технический скрипт: сохранить HTML отчёта (AnalyticsSpy.py)", "AnalyticsSpy.py"),
        },
    ),
}


def _sort_keys(keys):
    def _key(val: str):
        return (0, int(val)) if val.isdigit() else (1, val)

    return sorted(keys, key=_key)


def _run_script(title: str, rel_path: str) -> int:
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


def _menu_loop(
    options: Dict[str, Tuple[str, Union[str, Dict[str, Tuple[str, str]]]]],
    quit_label: str,
    allow_exit: bool,
) -> Optional[int]:
    keys_sorted = _sort_keys(options.keys())
    if not keys_sorted:
        print("[error] Нет доступных скриптов.")
        return 1
    choice_prompt = "Ваш выбор: "
    while True:
        print("Выберите скрипт для запуска:")
        for key in keys_sorted:
            title, target = options[key]
            suffix = " (подменю)" if isinstance(target, dict) else ""
            print(f" {key}. {title}{suffix}")
        print(f" q. {quit_label}")

        try:
            choice = input(choice_prompt).strip()
        except EOFError:
            choice = "q"
        if choice.lower() in {"q", "quit", "exit"}:
            return 0 if allow_exit else None
        if choice in options:
            title, target = options[choice]
            if isinstance(target, dict):
                result = _menu_loop(target, "Назад", allow_exit=False)
                if result is not None:
                    return result
                continue
            return _run_script(title, target)
        else:
            print("Введите номер пункта меню или q для выхода.")


def main() -> int:
    result = _menu_loop(SCRIPTS, "Выход", allow_exit=True)
    return result if result is not None else 0

if __name__ == "__main__":
    raise SystemExit(main())
