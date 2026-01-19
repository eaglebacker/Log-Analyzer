@echo off
echo Starting Log Analyzer...
python -m pip install openpyxl >nul 2>&1
python log_analyzer_gui.py
