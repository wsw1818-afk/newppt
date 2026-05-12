@echo off
chcp 65001 > nul
setlocal

echo DocumentExtractor v3 - security PC builds
echo.

echo Building one-file EXE without UPX...
py -m PyInstaller --clean --noconfirm DocumentExtractor_v3.spec
if errorlevel 1 goto :error

echo.
echo Building folder distribution without self-extract...
py -m PyInstaller --clean --noconfirm DocumentExtractor_v3_folder.spec
if errorlevel 1 goto :error

echo.
echo Copying outputs...
set "OUTPUT_DIR=D:\OneDrive\코드작업\결과물\newppt"
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"
copy /Y "dist\DocumentExtractor_v3.exe" "%OUTPUT_DIR%\"
if errorlevel 1 goto :error

if exist "%OUTPUT_DIR%\DocumentExtractor_v3_folder" rmdir /S /Q "%OUTPUT_DIR%\DocumentExtractor_v3_folder"
xcopy /E /I /Y "dist\DocumentExtractor_v3_folder" "%OUTPUT_DIR%\DocumentExtractor_v3_folder"
if errorlevel 1 goto :error

echo.
echo ========================================
echo Build Complete
echo One-file: %OUTPUT_DIR%\DocumentExtractor_v3.exe
echo Folder:   %OUTPUT_DIR%\DocumentExtractor_v3_folder\DocumentExtractor_v3.exe
echo ========================================
goto :done

:error
echo.
echo ========================================
echo Build Failed
echo ========================================
exit /b 1

:done
pause
