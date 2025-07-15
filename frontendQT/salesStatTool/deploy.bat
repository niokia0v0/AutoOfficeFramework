@echo off

echo --- [0/4] Cleaning up previous deployment...
if exist "%~dp0..\..\Releases\v0.1\salesStatTool.exe" (
    del /Q "%~dp0..\..\Releases\v0.1\*.*"
)
for %%d in (platforms styles iconengines imageformats translations) do (
    if exist "%~dp0..\..\Releases\v0.1\%%d" (
        rmdir /S /Q "%~dp0..\..\Releases\v0.1\%%d"
    )
)


echo --- [1/4] Copying backend engine...
robocopy "%~dp0..\..\Releases\v0.1\backend_engine" "%~dp0..\..\Releases\v0.1\backend_engine" /E /PURGE /NFL /NDL /NJH /NJS /nc /ns /np

echo --- [2/4] Copying frontend executable...
copy /Y "%~1\%~2.exe" "%~dp0..\..\Releases\v0.1"

echo --- [3/4] Deploying Qt libraries...
cd /d "%~dp0..\..\Releases\v0.1"
call "%~3\windeployqt.exe" %2.exe

echo --- [4/4] Deployment finished successfully!