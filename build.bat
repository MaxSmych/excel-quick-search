@echo off
rem Rebuild "Excel Search.exe" (onefile, own icon) from source.
rem Builds in %TEMP% (network drives break PyInstaller), then copies exe back.
setlocal
set "SRC=%~dp0"
set "B=%TEMP%\excel_search_build"

if exist "%B%" rmdir /s /q "%B%"
mkdir "%B%"
copy /Y "%SRC%excel_search.py" "%B%\" >nul
copy /Y "%SRC%icon.ico"        "%B%\" >nul

pushd "%B%"
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "Excel Search" --icon icon.ico --add-data "icon.ico;." excel_search.py
popd

copy /Y "%B%\dist\Excel Search.exe" "%SRC%" >nul && echo [OK] Excel Search.exe rebuilt
pause
