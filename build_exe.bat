@echo off
chcp 65001
echo 문서 보안 해제 저장 도구 - EXE 빌드
echo.

REM 패키지 설치
echo 패키지 설치 중...
pip install -r requirements.txt

echo.
echo EXE 빌드 중...
pyinstaller --onefile --windowed --name "DocumentExtractor" --icon=NONE document_extractor.py

echo.
echo 결과물 폴더로 복사 중...
set OUTPUT_DIR=D:\OneDrive\코드작업\결과물\DocumentExtractor
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"
copy /Y "dist\DocumentExtractor.exe" "%OUTPUT_DIR%\"
copy /Y "사용설명서.txt" "%OUTPUT_DIR%\"

echo.
echo ========================================
echo 빌드 완료!
echo 결과물 위치: %OUTPUT_DIR%
echo   - DocumentExtractor.exe
echo   - 사용설명서.txt
echo ========================================
pause
