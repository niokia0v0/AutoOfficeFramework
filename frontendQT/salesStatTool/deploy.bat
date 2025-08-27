@echo off
setlocal

:: --- [0] 从 .pro 文件获取参数 ---
set DEPLOY_DIR=%~1
set TARGET_NAME=%~2
set QT_BIN_PATH=%~3

echo [INFO] Starting deployment script...
echo [INFO] Target Dir : %DEPLOY_DIR%
echo [INFO] EXE Name   : %TARGET_NAME%.exe
echo [INFO] Qt Bin Path: %QT_BIN_PATH%
echo.

:: --- [1] 运行 windeployqt, 将所有依赖项复制到根目录 ---
echo [STEP 1/4] Running windeployqt to gather dependencies...
call "%QT_BIN_PATH%\windeployqt.exe" "%DEPLOY_DIR%\%TARGET_NAME%.exe"
if %errorlevel% neq 0 (
    echo [ERROR] windeployqt failed. Deployment aborted.
    exit /b 1
)

:: --- [2] 创建新的 frontend_runtime 目录 ---
set RUNTIME_DIR=%DEPLOY_DIR%\frontend_runtime
echo [STEP 2/4] Creating runtime directory...
if not exist "%RUNTIME_DIR%" mkdir "%RUNTIME_DIR%"

:: --- [3] 将 Qt 相关的依赖项移动到新目录 ---
echo [STEP 3/4] Moving Qt dependencies into runtime directory...

:: 移动 Qt 的 DLL 文件
:: 使用 > nul 来抑制 "1 file(s) moved." 的输出
move /Y "%DEPLOY_DIR%\Qt*.dll"           "%RUNTIME_DIR%\" > nul
move /Y "%DEPLOY_DIR%\libEGL.dll"        "%RUNTIME_DIR%\" > nul
move /Y "%DEPLOY_DIR%\libGLESv2.dll"     "%RUNTIME_DIR%\" > nul
move /Y "%DEPLOY_DIR%\opengl32sw.dll"    "%RUNTIME_DIR%\" > nul
move /Y "%DEPLOY_DIR%\D3DCompiler_47.dll" "%RUNTIME_DIR%\" > nul

:: 移动 Qt 的插件文件夹
:: 使用 if exist 避免因文件夹不存在而报错
if exist "%DEPLOY_DIR%\platforms"    move /Y "%DEPLOY_DIR%\platforms"    "%RUNTIME_DIR%\"
if exist "%DEPLOY_DIR%\styles"       move /Y "%DEPLOY_DIR%\styles"       "%RUNTIME_DIR%\"
if exist "%DEPLOY_DIR%\iconengines"  move /Y "%DEPLOY_DIR%\iconengines"  "%RUNTIME_DIR%\"
if exist "%DEPLOY_DIR%\imageformats" move /Y "%DEPLOY_DIR%\imageformats" "%RUNTIME_DIR%\"
if exist "%DEPLOY_DIR%\translations" move /Y "%DEPLOY_DIR%\translations" "%RUNTIME_DIR%\"

:: --- [4] 确保 MinGW C++ 核心运行库保留在根目录 ---
:: 这些库不是Qt库, 而是编译器运行库, 必须和exe在同一目录
echo [STEP 4/4] Ensuring MinGW runtime libraries exist in root...
copy /Y "%QT_BIN_PATH%\libgcc_s_seh-1.dll"    "%DEPLOY_DIR%\" > nul
copy /Y "%QT_BIN_PATH%\libstdc++-6.dll"     "%DEPLOY_DIR%\" > nul
copy /Y "%QT_BIN_PATH%\libwinpthread-1.dll" "%DEPLOY_DIR%\" > nul

echo.
echo [SUCCESS] Deployment finished successfully!

endlocal