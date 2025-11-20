@echo off
echo ============================================
echo   Kill All Processes and Restart
echo ============================================
echo.

echo Killing Overlay processes...
taskkill /F /IM WordOverlayProofreader.Overlay.exe 2>NUL
timeout /t 1 /nobreak > nul

echo Killing Addin test processes...
taskkill /F /IM WordOverlayProofreader.Addin.exe 2>NUL
timeout /t 1 /nobreak > nul

echo.
echo All processes killed. Now rebuilding...
dotnet build WordOverlayProofreader.sln --configuration Debug

echo.
echo ============================================
echo   Starting Fresh
echo ============================================
echo.
pause

echo Starting Overlay (you should see console output)...
start "Overlay Console" "src\WordOverlayProofreader.Overlay\bin\Debug\net8.0-windows\WordOverlayProofreader.Overlay.exe"
timeout /t 3 /nobreak > nul

echo.
echo Now run the test app:
echo   src\WordOverlayProofreader.Addin\bin\Debug\net48\WordOverlayProofreader.Addin.exe
echo.
pause
