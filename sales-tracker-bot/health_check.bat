@echo off
REM 检查 watchdog 是否存活，死了就拉起（watchdog 自身有 PID 锁，不会重复启动）
cd /d D:\Eleven-cc\sales-tracker-bot

REM 用心跳文件判断（watchdog 每 60 秒写一次）
set ALIVE=0
if exist .watchdog_heartbeat (
    powershell -Command "$ts=[int](Get-Content .watchdog_heartbeat); $age=(Get-Date)-([DateTime]::new(1970,1,1,0,0,0,0).AddSeconds($ts).ToLocalTime()); if ($age.TotalMinutes -lt 10) { exit 0 } else { exit 1 }" >nul 2>&1
    if %errorlevel% equ 0 set ALIVE=1
)

if %ALIVE% equ 0 (
    echo [%date% %time%] Watchdog not alive, starting...
    start "" /MIN C:\Users\huoli\AppData\Local\Programs\Python\Python312\python.exe -u watchdog.py >> watchdog.log 2>&1
)
