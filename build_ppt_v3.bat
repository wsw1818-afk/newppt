@echo off
echo PPT Extractor v3 - EXE Build
echo.

pip install pywin32 pyinstaller python-pptx

echo.
echo Building EXE...
pyinstaller --onefile --windowed --name PPTExtractor_v3 ppt_extractor_v3.py

echo.
echo Copying to output folder...
if not exist "D:\OneDrive\코드작업\결과물\DocumentExtractor" mkdir "D:\OneDrive\코드작업\결과물\DocumentExtractor"
copy /Y "dist\PPTExtractor_v3.exe" "D:\OneDrive\코드작업\결과물\DocumentExtractor\"

echo.
echo ========================================
echo Build Complete!
echo ========================================
pause
