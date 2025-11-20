@echo off
echo ============================================
echo   Word Overlay Proofreader - Quick Start
echo ============================================
echo.
echo This script will start all components for testing.
echo.
echo BEFORE RUNNING: Please open Microsoft Word!
echo.
pause

echo.
echo [1/2] Starting Overlay Application...
start "Overlay" "src\WordOverlayProofreader.Overlay\bin\Debug\net8.0-windows\WordOverlayProofreader.Overlay.exe"
timeout /t 2 /nobreak > nul

echo [2/2] Starting Add-in Test Application...
echo.
echo ============================================
echo   Instructions:
echo   - Press 's' to scan your document
echo   - Press 'a' to toggle auto-scan
echo   - Press 'q' to quit
echo ============================================
echo.
"src\WordOverlayProofreader.Addin\bin\Debug\net48\WordOverlayProofreader.Addin.exe"

echo.
echo Shutting down...
