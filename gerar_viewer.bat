@echo off
setlocal

if not exist .venv (
    py -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python src\build_viewer.py

start "" dist\index.html
endlocal
