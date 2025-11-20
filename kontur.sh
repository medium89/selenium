#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="${PROJECT_ROOT}/venv"
if [[ ! -d "${VENV_DIR}" ]]; then
    VENV_DIR="${PROJECT_ROOT}/.venv"
fi

if [[ ! -d "${VENV_DIR}" ]]; then
    echo "Виртуальное окружение venv не найдено по пути ${PROJECT_ROOT}/venv" >&2
    exit 1
fi

# shellcheck source=/dev/null
source "${VENV_DIR}/bin/activate"

export HEADLESS="${HEADLESS:-1}"

python "${PROJECT_ROOT}/app/kontur.py" "$@"
