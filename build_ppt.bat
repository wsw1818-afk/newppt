@echo off
echo PPT Extractor - EXE Build
echo.

pip install pywin32 pyinstaller

echo.
echo Building EXE...
pyinstaller --onefile --windowed --name PPTExtractor ppt_extractor.py

echo.
echo Copying to output folder...
if not exist "D:\OneDrive\코드작업\결과물\DocumentExtractor" mkdir "D:\OneDrive\코드작업\결과물\DocumentExtractor"
copy /Y "dist\PPTExtractor.exe" "D:\OneDrive\코드작업\결과물\DocumentExtractor\"

echo.
echo ========================================
echo Build Complete!
echo ========================================
pause
