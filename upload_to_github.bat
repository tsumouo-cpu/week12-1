@echo off
chcp 65001 >nul
echo ========================================================
echo     航運專案 - GitHub 一鍵自動上傳工具
echo ========================================================
echo 本地專案路徑: C:\proj\week12-1
echo 設定的電子郵件: sandy91011706@gmail.com
echo ========================================================
echo.

set /p GITHUB_USER="請輸入您的 GitHub 使用者名稱 (Username): "
if "%GITHUB_USER%"=="" (
    echo [錯誤] 使用者名稱不能為空！
    pause
    exit /b 1
)

set /p REPO_NAME="請輸入欲上傳的 Repository 名稱 [預設 week12-1]: "
if "%REPO_NAME%"=="" set REPO_NAME=week12-1

echo.
echo 即將連結至: https://github.com/%GITHUB_USER%/%REPO_NAME%.git
echo.

cd /d C:\proj\week12-1
"C:\Program Files\Git\cmd\git.exe" remote remove origin 2>nul
"C:\Program Files\Git\cmd\git.exe" remote add origin https://github.com/%GITHUB_USER%/%REPO_NAME%.git
"C:\Program Files\Git\cmd\git.exe" branch -M main

echo 正在推送至 GitHub... (稍後若跳出瀏覽器或登入視窗，請點選登入授權)
"C:\Program Files\Git\cmd\git.exe" push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================================
    echo [成功] 專案已成功上傳至: https://github.com/%GITHUB_USER%/%REPO_NAME%
    echo ========================================================
) else (
    echo.
    echo ========================================================
    echo [提示] 若上傳失敗，通常是因為:
    echo 1. 您尚未在 GitHub 網頁上建立名為 "%REPO_NAME%" 的 Repository。
    echo    請先至 https://github.com/new 建立該儲存庫後再次執行。
    echo 2. 帳號或密碼/Token 驗證未通過。
    echo ========================================================
)

pause
