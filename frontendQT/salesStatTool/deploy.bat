@echo off
setlocal

:: 接收来自 .pro 文件的参数
:: %1: 部署目录 (DESTDIR, e.g., "F:\projects\AutoOfficeFramework\Releases\salesStatTool")
:: %2: 目标程序名 (TARGET, e.g., "salesStatTool")
:: %3: Qt的bin目录 (e.g., "C:\Qt\5.14.2\mingw73_64\bin")

set DEPLOY_DIR=%~1
set TARGET_NAME=%~2
set QT_BIN_PATH=%~3

echo --- [1/2] Entering deployment directory...
echo      %DEPLOY_DIR%
cd /d "%DEPLOY_DIR%"

if not exist "%TARGET_NAME%.exe" (
    echo ERROR: Target executable '%TARGET_NAME%.exe' not found in deployment directory!
    exit /b 1
)

echo --- [2/2] Deploying Qt libraries using windeployqt...
call "%QT_BIN_PATH%\windeployqt.exe" "%TARGET_NAME%.exe"

echo --- Deployment finished successfully! ---

endlocal