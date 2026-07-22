@echo off
rem Rebuild "Excel Search" (onedir, own icon) from source.
rem onedir = fast startup (no per-run unpack). Builds in %TEMP% (network drives break
rem PyInstaller), then zips the whole app folder back (AV deletes bare exe on the share).
setlocal
set "SRC=%~dp0"
set "B=%TEMP%\excel_search_build"

if exist "%B%" rmdir /s /q "%B%"
mkdir "%B%"
copy /Y "%SRC%excel_search.py" "%B%\" >nul
copy /Y "%SRC%icon.ico"        "%B%\" >nul

pushd "%B%"
python -m PyInstaller --noconfirm --clean --onedir --windowed ^
    --name "Excel Search" --icon icon.ico --add-data "icon.ico;." excel_search.py
popd

rem Pack the CONTENTS of the app folder (exe + _internal) so they extract straight
rem into the target folder next to excel_search_settings.json (shared with the .bat).
if exist "%SRC%Excel Search.zip" del "%SRC%Excel Search.zip"
tar -a -c -f "%SRC%Excel Search.zip" -C "%B%\dist\Excel Search" . && echo [OK] Excel Search.zip rebuilt
pause
