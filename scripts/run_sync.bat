@echo off
setlocal

cd /d "%~dp0.."

echo Running journal sync...
call .venv\Scripts\journal-tracker-sync --profile config/profiles/turkish-politics-starter.json --workbook data\turkish_politics_articles_database.xlsx

echo.
echo Done. Press any key to close.
pause >nul
