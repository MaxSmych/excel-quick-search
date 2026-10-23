@echo off
rem Installer "Excel-Search-2-Setup.exe": one file with Excel-Search-Qt.zip inside (run build_qt.bat first).
setlocal
set "SRC=%~dp0"
set "B=%TEMP%\excel_search_setup_build"
if not exist "%SRC%Excel-Search-Qt.zip" (echo [ERR] run build_qt.bat first & exit /b 1)

if exist "%B%" rmdir /s /q "%B%"
mkdir "%B%"
copy /Y "%SRC%setup_qt.py"         "%B%\" >nul
copy /Y "%SRC%icon.ico"            "%B%\" >nul
copy /Y "%SRC%Excel-Search-Qt.zip" "%B%\" >nul

pushd "%B%"
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "Excel-Search-2-Setup" --icon icon.ico --add-data "Excel-Search-Qt.zip;." setup_qt.py
popd

copy /Y "%B%\dist\Excel-Search-2-Setup.exe" "%SRC%" >nul && echo [OK] Excel-Search-2-Setup.exe
