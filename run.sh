#!/usr/bin/env bash
# Kleiner Wrapper: nutzt die mitgelieferte virtuelle Umgebung (.venv).
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
exec "$here/.venv/bin/python" "$here/einreichen.py" "$@"
