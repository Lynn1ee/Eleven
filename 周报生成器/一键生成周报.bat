@echo off
:: 客服数据周报一键生成 —— 把 Excel 文件拖到这个 bat 图标上即可
:: 输出到 Excel 同目录，HTML + PNG 各一份

if "%~1"=="" (
    echo 用法：把 Excel 文件拖到这个 bat 图标上
    pause
    exit /b
)

set SCRIPT_DIR=%~dp0
set EXCEL_PATH=%~1
set DIR=%~dp1
set NAME=%~n1
set LOG_FILE=%DIR%%NAME%_生成日志.txt

echo [%date% %time%] 开始生成周报... > "%LOG_FILE%"
echo   源文件: %EXCEL_PATH% >> "%LOG_FILE%"

echo.
echo [1/2] 生成 HTML 周报...

python "%SCRIPT_DIR%gen_weekly_report.py" "%EXCEL_PATH%" "%DIR%%NAME%.html" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo HTML 生成失败！详情见日志文件：
    echo   %LOG_FILE%
    type "%LOG_FILE%"
    pause
    exit /b
)
echo   HTML 已生成 >> "%LOG_FILE%"

echo [2/2] 生成 PNG 截图...
python "%SCRIPT_DIR%screenshot_report.py" "%DIR%%NAME%.html" "%DIR%%NAME%.png" >> "%LOG_FILE%" 2>&1
if %errorlevel% neq 0 (
    echo PNG 截图失败！详情见日志文件：
    echo   %LOG_FILE%
    type "%LOG_FILE%"
    pause
    exit /b
)
echo   PNG 已生成 >> "%LOG_FILE%"

echo.
echo ============================
echo 完成！输出文件：
echo   %DIR%%NAME%.html
echo   %DIR%%NAME%.png
echo ============================
echo 完成 >> "%LOG_FILE%"
start "" "%DIR%%NAME%.html"
pause
