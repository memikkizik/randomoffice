#!/usr/bin/env bash
set -e

HOST="${HOST:-0.0.0.0}"
PORT="${PORT:-8000}"

echo "=========================================="
echo " Starting RandomOffice Enterprise Service"
echo "=========================================="

if [ -f ".venv/bin/uvicorn" ]; then
    echo "Starting Uvicorn server on http://${HOST}:${PORT}..."
    exec .venv/bin/uvicorn randomoffice.api.app:app --host "${HOST}" --port "${PORT}"
else
    echo "Using uv run..."
    exec uv run uvicorn randomoffice.api.app:app --host "${HOST}" --port "${PORT}"
fi

