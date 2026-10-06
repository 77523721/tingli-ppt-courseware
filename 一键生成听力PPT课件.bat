@echo off
rem One-click build of "<basename>听力课件.pptx" in this folder.
rem Run with no args = build here; pass a folder path to build that folder instead.
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
python "%~dp0make_ppt_courseware.py" %*
echo.
pause
