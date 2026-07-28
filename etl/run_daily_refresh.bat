@echo off
cd /d "%~dp0.."
echo [%date% %time%] ---- daily refresh start ---- >> etl\refresh.log
C:\Python314\python.exe etl\load_excel_to_db.py >> etl\refresh.log 2>&1
echo [%date% %time%] ---- daily refresh end ---- >> etl\refresh.log
