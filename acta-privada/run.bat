@echo off
cd /d "%~dp0"
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
python -m streamlit run app.py --server.address 127.0.0.1 --server.headless true
