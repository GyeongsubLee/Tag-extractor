@echo off
REM Drop this file in the PDF folder (or pass a folder) and double-click
python "%~dp0valve_tag_extractor.py" %* "%CD%"
pause
