@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Deploy to Google Cloud Run - TT27 Assessor AI

echo ========================================================
echo   TRIEN KHAI TT27 ASSESSOR AI LEN GOOGLE CLOUD RUN
echo ========================================================
echo.

where gcloud >nul 2>&1
if %errorlevel% neq 0 (
    echo [THONG BAO] Chua tim thay gcloud CLI tren may tinh.
    echo.
    echo Ban co the:
    echo 1. Cai dat Google Cloud SDK: https://cloud.google.com/sdk/docs/install
    echo 2. HOAC su dung Google Cloud Shell tren trinh duyet (khong can cai dat).
    echo.
    pause
    exit /b 1
)

set /p PROJECT_ID="Nhap Google Cloud Project ID cua ban: "
if "%PROJECT_ID%"=="" (
    echo [LOI] Project ID khong duoc de trong!
    pause
    exit /b 1
)

echo.
echo [*] Dang thiet lap project: %PROJECT_ID%...
gcloud config set project %PROJECT_ID%

echo.
echo [*] Dang bat cac API can thiet tren Google Cloud...
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com

echo.
echo [*] Dang tien hanh deploy len Google Cloud Run...
gcloud run deploy tt27-assessor-ai ^
  --source . ^
  --region asia-southeast1 ^
  --allow-unauthenticated ^
  --memory 1Gi ^
  --cpu 1 ^
  --timeout 300s ^
  --session-affinity

echo.
echo ========================================================
echo Hoan tat qua trinh trien khai!
echo ========================================================
pause
