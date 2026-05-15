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
    echo Python 3 설치 후 다시 실행하세요.
    pause
    exit /b 1
)

%PYTHON_CMD% -m pip install --upgrade pip
if errorlevel 1 goto :error

%PYTHON_CMD% -m pip install -r "%~dp0requirements_runtime.txt"
if errorlevel 1 goto :error

echo.
echo 설치 완료. run_from_source.bat를 실행하세요.
pause
exit /b 0

:error
echo.
echo 패키지 설치 실패.
echo 회사망에서 인터넷/pip가 막힌 경우 IT 보안 예외 또는 내부 PyPI 설정이 필요합니다.
pause
exit /b 1
