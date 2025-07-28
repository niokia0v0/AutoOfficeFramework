@echo off

echo --- [1/4] Cleaning up previous deployment...
if exist "%~dp0..\..\Releases\salesStatTool\salesStatTool.exe" (
    del /Q "%~dp0..\..\Releases\salesStatTool\*.*"
)
for %%d in (platforms styles iconengines imageformats translations) do (
    if exist "%~dp0..\..\Releases\salesStatTool\%%d" (
        rmdir /S /Q "%~dp0..\..\Releases\salesStatTool\%%d"
    )
)

echo --- [2/4] Copying frontend executable...
copy /Y "%~1\%~2.exe" "%~dp0..\..\Releases\salesStatTool"

echo --- [3/4] Deploying Qt libraries...
cd /d "%~dp0..\..\Releases\salesStatTool"
call "%~3\windeployqt.exe" %2.exe

echo --- [4/4] Deployment finished successfully!