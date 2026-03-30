@echo off
setlocal
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set BUILD_TS=%%i
set EXE_NAME=Logbooker_%BUILD_TS%

echo [1/3] Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 goto :error

echo [2/3] Installing PyInstaller...
python -m pip install pyinstaller
if errorlevel 1 goto :error

echo [2.0/3] Installing Pillow for icon conversion...
python -m pip install pillow
if errorlevel 1 goto :error

if exist iconMY.png (
    echo [2.1/3] Converting iconMY.png to iconMY.ico...
    python convert_icon.py
    if errorlevel 1 goto :error
)

if exist iconMY.ico (
    set ICON_ARG=--icon "%CD%\iconMY.ico"
) else if exist app_icon.ico (
    set ICON_ARG=--icon "%CD%\app_icon.ico"
) else (
    set ICON_ARG=
)

rem Ресурсы для runtime (иконка в заголовке и на панели задач): без --add-data файлов нет в _MEIPASS.
set ADD_DATA=
if exist iconMY.png set ADD_DATA=%ADD_DATA% --add-data "iconMY.png;."
if exist iconMY.ico set ADD_DATA=%ADD_DATA% --add-data "iconMY.ico;."
if exist app_icon.png set ADD_DATA=%ADD_DATA% --add-data "app_icon.png;."
if exist app_icon.ico set ADD_DATA=%ADD_DATA% --add-data "app_icon.ico;."

echo [3/3] Building EXE...
python -m PyInstaller --noconfirm --clean --windowed --onefile --name %EXE_NAME% %ICON_ARG% %ADD_DATA% app.py
if errorlevel 1 goto :error

echo.
echo Build complete.
echo EXE: dist\%EXE_NAME%.exe
exit /b 0

:error
echo.
echo Build failed. Check messages above.
exit /b 1
