@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 ( py recompra_app.py ) else ( python recompra_app.py )
if %errorlevel% neq 0 (
    echo.
    echo Ocorreu um erro. Verifique se o Python esta instalado e se rodou:
    echo     pip install -r requirements.txt
    echo.
    echo Dependencias: openpyxl e, para a fonte "Banco de dados", psycopg2-binary.
    echo.
    pause
)
