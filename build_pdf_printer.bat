@echo off
chcp 65001
echo 문서 PDF 인쇄 도구 - EXE 빌드
echo.

REM 패키지 설치
echo 패키지 설치 중...
pip install pywin32 pyinstaller

echo.
echo EXE 빌드 중...
pyinstaller --onefile --windowed --name "PDFPrinter" pdf_printer.py

echo.
echo 결과물 폴더로 복사 중...
set OUTPUT_DIR=D:\OneDrive\코드작업\결과물\DocumentExtractor
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"
copy /Y "dist\PDFPrinter.exe" "%OUTPUT_DIR%\"
copy /Y "PDF인쇄_사용설명서.txt" "%OUTPUT_DIR%\"

echo.
echo ========================================
echo 빌드 완료!
echo 결과물 위치: %OUTPUT_DIR%
echo   - PDFPrinter.exe
echo   - PDF인쇄_사용설명서.txt
echo ========================================
pause
