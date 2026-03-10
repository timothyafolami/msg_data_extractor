#!/usr/bin/env bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
VENV_PYTHON="$SCRIPT_DIR/.venv/bin/python"

usage() {
  cat <<'EOF'
Usage:
  ./process_folder.sh <input-folder> [output-folder]

Examples:
  ./process_folder.sh /path/to/client_folder
  ./process_folder.sh /path/to/client_folder /path/to/output_folder
EOF
}

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

if [[ $# -lt 1 || $# -gt 2 ]]; then
  usage
  exit 1
fi

if [[ ! -x "$VENV_PYTHON" ]]; then
  echo "Virtual environment not found. Run ./setup_local.sh first." >&2
  exit 1
fi

INPUT_FOLDER="$1"
OUTPUT_FOLDER="${2:-}"

if [[ ! -d "$INPUT_FOLDER" ]]; then
  echo "Input folder does not exist or is not a directory: $INPUT_FOLDER" >&2
  exit 1
fi

if [[ -z "$OUTPUT_FOLDER" ]]; then
  INPUT_BASENAME="$(basename "$INPUT_FOLDER")"
  OUTPUT_FOLDER="$SCRIPT_DIR/${INPUT_BASENAME}_extracted"
fi

WORKERS="${WORKERS:-}"

if [[ -n "$WORKERS" ]]; then
  exec "$VENV_PYTHON" "$SCRIPT_DIR/extract_msg_photos.py" "$INPUT_FOLDER" --output-folder "$OUTPUT_FOLDER" --workers "$WORKERS"
fi

exec "$VENV_PYTHON" "$SCRIPT_DIR/extract_msg_photos.py" "$INPUT_FOLDER" --output-folder "$OUTPUT_FOLDER"