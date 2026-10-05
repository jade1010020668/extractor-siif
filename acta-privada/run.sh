#!/usr/bin/env bash
# Inicia la app SOLO en este equipo (127.0.0.1), sin telemetría.
cd "$(dirname "$0")"
export STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
exec python3 -m streamlit run app.py --server.address 127.0.0.1 --server.headless true
