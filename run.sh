#!/usr/bin/env bash
# Linux / macOS: reads every PDF in data/ and writes result/yields.xlsx and result/yields.csv
cd "$(dirname "$0")" || exit 1
if command -v python3 >/dev/null 2>&1; then PY=python3; else PY=python; fi
"$PY" extract_yields.py "$@"
