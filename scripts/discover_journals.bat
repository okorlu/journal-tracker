@echo off
setlocal

cd /d "%~dp0.."

echo Running journal discovery...
call .venv\Scripts\journal-tracker-discover-journals --workbook data\turkish_politics_articles_database.xlsx

echo.
echo Done. Press any key to close.
pause >nul
