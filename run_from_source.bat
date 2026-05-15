@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

set "PYTHON_CMD="
where py > nul 2> nul
if not errorlevel 1 set "PYTHON_CMD=py -3"

if not defined PYTHON_CMD (
    where python > nul 2> nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
    echo Python을 찾을 수 없습니다.
    echo 회사 PC에서 EXE가 차단되면 Python 3 설치 후 다시 실행하세요.
    echo 설치 후 install_requirements.bat를 먼저 실행해야 합니다.
    pause
    exit /b 1
)

%PYTHON_CMD% -c "import win32com.client, pythoncom, pptx, openpyxl, docx, PIL, lxml" > nul 2> nul
if errorlevel 1 (
    echo 필요한 Python 패키지가 없습니다.
    echo install_requirements.bat를 먼저 실행하세요.
    pause
    exit /b 1
)

%PYTHON_CMD% "%~dp0ppt_extractor_v3.py"
if errorlevel 1 (
    echo.
    echo 프로그램 실행 중 오류가 발생했습니다.
    echo 바탕화면의 DocExtractor_Startup_Error_*.txt 또는 DocExtractor_Log_*.txt를 확인하세요.
    pause
    exit /b 1
)
