@echo off
title Paltrack Masterfile Builder
set "PY=C:\Users\Kamelihled\AppData\Local\Programs\Python\Python313\python.exe"
if not exist "%PY%" set "PY=py"

rem Install the two packages the first time only
"%PY%" -c "import xlsxwriter, openpyxl" 2>nul || "%PY%" -m pip install xlsxwriter openpyxl

"%PY%" "%~dp0paltrack_masterfile_builder.py"
echo.
pause
