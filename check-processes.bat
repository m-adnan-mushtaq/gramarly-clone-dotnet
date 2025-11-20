@echo off
echo ============================================
echo   Checking Running Processes
echo ============================================
echo.

echo Checking for Overlay application...
tasklist /FI "IMAGENAME eq WordOverlayProofreader.Overlay.exe" 2>NUL | find /I /N "WordOverlayProofreader.Overlay.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo [OK] Overlay is RUNNING
) else (
    echo [!!] Overlay is NOT running
    echo     Start it with: src\WordOverlayProofreader.Overlay\bin\Debug\net8.0-windows\WordOverlayProofreader.Overlay.exe
)

echo.
echo Checking for Word...
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I /N "WINWORD.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo [OK] Word is RUNNING
) else (
    echo [!!] Word is NOT running
    echo     Please open Microsoft Word
)

echo.
echo ============================================
pause
