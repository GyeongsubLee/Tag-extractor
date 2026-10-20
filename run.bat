@echo off
REM Run from the PDF folder (double-click), or: run.bat "D:\PID\PDFs"
if "%~1"=="" (python "%~dp0valve_tag_extractor.py" "%CD%") else (python "%~dp0valve_tag_extractor.py" %*)
pause
