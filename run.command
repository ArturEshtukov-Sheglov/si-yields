#!/usr/bin/env bash
# macOS: double-click in Finder (runs run.sh in Terminal)
cd "$(dirname "$0")" || exit 1
bash run.sh
echo
read -n 1 -s -r -p "Press any key to close"
