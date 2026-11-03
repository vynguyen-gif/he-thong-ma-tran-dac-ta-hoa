@echo off
title KHOI TAO VA DAY DU AN LEN GITHUB
chcp 65001 >nul
cd /d "%~dp0"

set "GIT_CMD=git"
where git >nul 2>&1
if %errorlevel% neq 0 (
    if exist "C:\Program Files\Git\cmd\git.exe" (
        set "GIT_CMD=C:\Program Files\Git\cmd\git.exe"
    ) else (
        echo [THONG BAO] May tinh cua ban chua co Git!
        echo Vui long cai dat Git tai https://git-scm.com/download/win
        pause
        exit /b
    )
)

echo =======================================================
echo    KHOI TAO VA DAY CODE LEN GITHUB (vynguyen-gif)
echo =======================================================
echo.
echo * Luu y: Truoc khi chay file nay, ban hay vao https://github.com/new
echo   de tao mot Repository moi (vi du dat ten la: he-thong-ma-tran-dac-ta-hoa).
echo.
set /p repoName="Nhap ten Repository tren GitHub (Enter de lay mac dinh: he-thong-ma-tran-dac-ta-hoa): "
if "%repoName%"=="" set repoName=he-thong-ma-tran-dac-ta-hoa

echo.
echo [1/4] Dang kiem tra Git kho luu tru...
if not exist ".git" (
    "%GIT_CMD%" init
)

echo [2/4] Dang chuan bi toan bo ma nguon...
"%GIT_CMD%" add .
"%GIT_CMD%" commit -m "Khoi tao he thong ma tran va ban dac ta de kiem tra Hoa hoc THPT CV 7991" 2>nul
"%GIT_CMD%" branch -M main

echo [3/4] Dang ket noi toi GitHub cua ban (https://github.com/vynguyen-gif/%repoName%.git)...
"%GIT_CMD%" remote remove origin >nul 2>&1
"%GIT_CMD%" remote add origin https://github.com/vynguyen-gif/%repoName%.git

echo [4/4] Dang day code len GitHub...
echo * Neu trinh duyet mo cua so dang nhap GitHub, ban chi can chon "Sign in with your browser" de xac thuc.
"%GIT_CMD%" push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo =======================================================
    echo [HOAN TAT] Du an da duoc day len GitHub thanh cong!
    echo.
    echo BUOC TIEP THEO DE KET NOI VERCEL (CHI CAN LAM 1 LAN):
    echo 1. Truy cap https://vercel.com/new
    echo 2. Tim kho "%repoName%" va bam nut "Import"
    echo 3. Bam nut "Deploy" (he thong tinh toan tu dong 100%%)
    echo 4. Tu nay ve sau, moi khi sua code, ban chi can chay file
    echo    "DONG_BO_GITHUB_VERCEL.bat" la Vercel se tu dong cap nhat!
    echo =======================================================
) else (
    echo.
    echo [CANH BAO] Khong the day code len GitHub.
    echo Nguyen nhan thuong gap:
    echo 1. Ban chua tao repository "%repoName%" tren https://github.com/new
    echo 2. Hoac ten repository bi go sai.
    echo Hay tao repository tren GitHub roi chay lai file nay nhe!
)
echo.
pause
