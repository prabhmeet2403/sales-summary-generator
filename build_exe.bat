@echo off
REM build_exe.bat
REM ==============
REM Convenience script to build the standalone NVISH Sales Forecast
REM Automation.exe on Windows. Run from the project root:
REM
REM     build_exe.bat
REM
REM The finished executable is written to dist\NVISH Sales Forecast Automation.exe

setlocal

echo ============================================================
echo  NVISH Sales Forecast Automation - Windows Build
echo ============================================================

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found on PATH. Install Python 3.10+ and try again.
    exit /b 1
)

echo.
echo [1/3] Installing/upgrading dependencies...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    exit /b 1
)

echo.
echo [2/3] Regenerating brand assets (logo/icon)...
python gui\assets\generate_assets.py
if errorlevel 1 (
    echo WARNING: Asset generation failed; continuing with existing assets if present.
)

echo.
echo [3/3] Building the executable with PyInstaller...
pyinstaller --noconfirm --clean SalesForecastGUI.spec
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete.
echo  Executable: dist\NVISH Sales Forecast Automation.exe
echo ============================================================

endlocal
