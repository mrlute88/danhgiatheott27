@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"
title TT27 Assessor AI

set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo ========================================================
echo   TT27 Assessor AI - Tro ly Nhan xet Hoc sinh Tieu hoc
echo             Dang khoi dong ung dung web...
echo ========================================================
echo.

REM 1. Bo qua man hinh hoi email cua Streamlit
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit" 2>nul
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    (echo [general] & echo email = "") > "%USERPROFILE%\.streamlit\credentials.toml" 2>nul
)
if not exist ".streamlit" mkdir ".streamlit" 2>nul
if not exist ".streamlit\credentials.toml" (
    (echo [general] & echo email = "") > ".streamlit\credentials.toml" 2>nul
)

REM 2. Kich hoat moi truong ao neu co
if exist "venv\Scripts\activate.bat" (
    echo [*] Kich hoat moi truong ao venv...
    call "venv\Scripts\activate.bat"
) else if exist ".venv\Scripts\activate.bat" (
    echo [*] Kich hoat moi truong ao .venv...
    call ".venv\Scripts\activate.bat"
)

REM 3. Xac dinh duong dan Python
set "PY_CMD="
where python >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_CMD=python"
) else (
    where py >nul 2>&1
    if !errorlevel! equ 0 (
        set "PY_CMD=py"
    ) else if exist "C:\Python314\python.exe" (
        set "PY_CMD=C:\Python314\python.exe"
    )
)

if "%PY_CMD%"=="" (
    echo [LOI] Khong tim thay Python tren may tinh!
    echo Vui long cai dat Python 3.10 tro len tai: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 4. Kiem tra thu vien can thiet
echo [1/2] Kiem tra thu vien...
%PY_CMD% -c "import streamlit" >nul 2>&1
if !errorlevel! neq 0 (
    echo Dang cai dat thu vien tu requirements.txt, vui long doi...
    %PY_CMD% -m pip install -r requirements.txt
)

REM 5. Khoi chay Streamlit Web App
echo [2/2] Dang khoi chay Streamlit...
echo ========================================================
echo   Dia chi truy cap web:  http://localhost:8501
echo   De tat ung dung: Dong cua so nay hoac nhan Ctrl+C
echo ========================================================
echo.

%PY_CMD% -m streamlit run app.py --server.port 8501 --server.headless false

pause
