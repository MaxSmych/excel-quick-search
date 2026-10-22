@echo off
rem Rebuild "Excel Search 2" (Qt, onedir) from excel_search_qt.py.
rem Same scheme as build.bat: build in %TEMP%, zip the app folder back as Excel-Search-Qt.zip
rem (release asset; 1.x looks for Excel-Search.zip and never picks this one up).
setlocal
set "SRC=%~dp0"
set "B=%TEMP%\excel_search_qt_build"

if exist "%B%" rmdir /s /q "%B%"
mkdir "%B%"
copy /Y "%SRC%excel_search_qt.py" "%B%\" >nul
copy /Y "%SRC%icon.ico"           "%B%\" >nul
xcopy /E /I /Q /Y "%SRC%assets" "%B%\assets" >nul

pushd "%B%"
python -m PyInstaller --noconfirm --clean --onedir --windowed ^
    --name "Excel Search 2" --icon icon.ico --add-data "icon.ico;." ^
    --add-data "assets/splash;assets/splash" ^
    --hidden-import xlwings --hidden-import win32com.client --hidden-import pythoncom ^
    excel_search_qt.py
popd

if exist "%SRC%Excel-Search-Qt.zip" del "%SRC%Excel-Search-Qt.zip"
tar -a -c -f "%SRC%Excel-Search-Qt.zip" -C "%B%\dist\Excel Search 2" . && echo [OK] Excel-Search-Qt.zip rebuilt
pause
