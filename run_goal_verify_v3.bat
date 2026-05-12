@echo off
chcp 65001 > nul
setlocal

echo DocumentExtractor v3 - Goal Verification
echo.

py scripts\goal_verify_v3.py --clean
if errorlevel 1 goto :error

echo.
echo ========================================
echo Goal Verification Complete
echo ========================================
goto :done

:error
echo.
echo ========================================
echo Goal Verification Failed
echo ========================================
exit /b 1

:done
pause
