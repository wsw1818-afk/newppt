@echo off
chcp 65001 > nul
setlocal

echo DocumentExtractor v3 - EXE Build
echo.

echo Installing packages from requirements.txt...
py -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo Building EXE from DocumentExtractor_v3.spec...
py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec
if errorlevel 1 goto :error

echo.
echo Copying to output folder...
set "OUTPUT_DIR=D:\OneDrive\코드작업\결과물\newppt"
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"
copy /Y "dist\DocumentExtractor_v3.exe" "%OUTPUT_DIR%\"
if errorlevel 1 goto :error

echo.
echo ========================================
echo Build Complete!
echo Output: %OUTPUT_DIR%\DocumentExtractor_v3.exe
echo ========================================
goto :done

:error
echo.
echo ========================================
echo Build Failed!
echo ========================================
exit /b 1

:done
pause
