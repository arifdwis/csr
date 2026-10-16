#!/usr/bin/env bash
# Script peluncur dashboard analisis eksekutif CSR Samarinda
cd "$(dirname "$0")"

if [ -d ".venv" ]; then
    source .venv/bin/activate
fi

python3 viewer/dashboard_server.py
