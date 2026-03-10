#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
VENV_DIR="$SCRIPT_DIR/.venv"

if command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="$(command -v python3)"
elif command -v python >/dev/null 2>&1; then
  PYTHON_BIN="$(command -v python)"
else
  echo "Python 3 is required but was not found on PATH." >&2
  exit 1
fi

echo "Using Python: $PYTHON_BIN"

if [[ ! -d "$VENV_DIR" ]]; then
  echo "Creating virtual environment at $VENV_DIR"
  "$PYTHON_BIN" -m venv "$VENV_DIR"
else
  echo "Virtual environment already exists at $VENV_DIR"
fi

VENV_PYTHON="$VENV_DIR/bin/python"

echo "Upgrading pip"
"$VENV_PYTHON" -m pip install --upgrade pip

echo "Installing dependencies"
"$VENV_PYTHON" -m pip install -r "$SCRIPT_DIR/requirements.txt"

echo
echo "Setup complete."
echo "Run processing with:"
echo "  ./process_folder.sh /path/to/input_folder [/path/to/output_folder]"