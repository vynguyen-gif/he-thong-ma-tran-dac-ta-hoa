@echo off
title DONG BO GITHUB VA VERCEL - MA TRAN DAC TA HOA HOC
chcp 65001 >nul
cd /d "%~dp0"

set "GIT_CMD=git"
where git >nul 2>&1
if %errorlevel% neq 0 (
    if exist "C:\Program Files\Git\cmd\git.exe" (
        set "GIT_CMD=C:\Program Files\Git\cmd\git.exe"
    ) else (
        echo [CANH BAO] Khong tim thay Git tren may tinh!
        pause
        exit /b
    )
)

echo =======================================================
echo    DONG BO MA NGUON LEN GITHUB VA TU DONG DEPLOY VERCEL
echo =======================================================
echo.

echo [1/3] Kiem tra cac tap tin thay doi:
"%GIT_CMD%" status -s

echo.
set /p commitMsg="Nhap mo ta cap nhat (Bam Enter de dung mac dinh: 'Cap nhat he thong ma tran dac ta'): "
if "%commitMsg%"=="" set commitMsg=Cap nhat he thong ma tran dac ta

echo.
echo [2/3] Dang luu thay doi (Git Commit)...
"%GIT_CMD%" add .
"%GIT_CMD%" commit -m "%commitMsg%"

echo.
echo [3/3] Dang day code len GitHub (Git Push)...
"%GIT_CMD%" push origin main 2>nul
if %errorlevel% neq 0 (
    "%GIT_CMD%" push origin master
)

echo.
echo =======================================================
echo [HOAN TAT] Code da duoc day len GitHub thanh cong!
echo Vercel se TU DONG build va cap nhat trang web sau 30-60 giay.
echo =======================================================
echo.
pause
