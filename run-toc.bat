@echo off
chcp 65001 >nul
python "%~dp0bookmarks.py" "%~1" --toc-pages
pause
