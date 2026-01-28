#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REQ_FILE="${PROJECT_ROOT}/config/requirements.txt"
ENV_FILE="${PROJECT_ROOT}/.env"

if [[ -f "${ENV_FILE}" ]]; then
    set -a
    # shellcheck source=/dev/null
    source "${ENV_FILE}"
    set +a
fi

PY_BIN="${PY_BIN:-}"
if [[ -z "${PY_BIN}" ]]; then
    if command -v python3 >/dev/null 2>&1; then
        PY_BIN="python3"
    else
        PY_BIN="python"
    fi
fi

VENV_DIR="${VENV_DIR:-${PROJECT_ROOT}/venv}"
if [[ ! -d "${VENV_DIR}" ]]; then
    VENV_DIR="${PROJECT_ROOT}/.venv"
fi

VENV_CREATED="0"
if [[ -d "${VENV_DIR}" ]] && [[ ! -f "${VENV_DIR}/bin/activate" ]]; then
    TS="$(date +"%Y%m%d-%H%M%S" 2>/dev/null || echo "broken")"
    BROKEN_DIR="${VENV_DIR}.broken-${TS}"
    echo "[kontur] Найдено повреждённое venv (нет bin/activate), переименовываю в: ${BROKEN_DIR}" >&2
    mv "${VENV_DIR}" "${BROKEN_DIR}"
fi
if [[ ! -d "${VENV_DIR}" ]]; then
    echo "[kontur] venv не найден, создаю: ${VENV_DIR}" >&2
    if [[ ! -f "${REQ_FILE}" ]]; then
        echo "[kontur] Не найден файл зависимостей: ${REQ_FILE}" >&2
        exit 1
    fi
    if ! "${PY_BIN}" -m venv "${VENV_DIR}"; then
        PY_VER="$("${PY_BIN}" -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")' 2>/dev/null || true)"
        echo "[kontur] Не удалось создать venv через ensurepip." >&2
        if command -v apt-get >/dev/null 2>&1; then
            if [[ -n "${PY_VER}" ]]; then
                echo "[kontur] Для Debian/Ubuntu обычно нужно: sudo apt install -y python${PY_VER}-venv" >&2
            fi
            echo "[kontur] Альтернатива: sudo apt install -y python3-venv" >&2
        else
            echo "[kontur] Установите пакет venv для вашей ОС (python venv/ensurepip), затем повторите запуск." >&2
        fi
        exit 1
    fi
    VENV_CREATED="1"
fi

# shellcheck source=/dev/null
source "${VENV_DIR}/bin/activate"

export HEADLESS="${HEADLESS:-1}"

if [[ "${VENV_CREATED}" == "1" || "${KONTUR_PIP_INSTALL:-0}" == "1" ]] && [[ -f "${REQ_FILE}" ]]; then
    python -m pip install -r "${REQ_FILE}" >/dev/null
fi

python "${PROJECT_ROOT}/app/kontur.py" "$@"
