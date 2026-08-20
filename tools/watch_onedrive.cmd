@echo off
rem Starts the OneDrive -> GitHub Pages sync watcher. Leave this window open.
cd /d "%~dp0.."
python tools\watch_onedrive.py
pause
