@echo off
REM 东方财富 财报/研报 批量下载器 —— Windows 启动脚本
REM 用法： run.bat --code 600519 --period 2026 --zip
REM        run.bat --name 五粮液 --period 2026H1 --types report
REM        run.bat --doctor
setlocal
set "DIR=%~dp0"
set "APP=%DIR%eastmoney_downloader.py"

if not exist "%APP%" (
  echo [错误] 未找到 eastmoney_downloader.py，请确认本脚本与程序在同一目录。
  pause
  exit /b 1
)

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "%APP%" %*
) else (
  where python >nul 2>nul
  if %errorlevel%==0 (
    python "%APP%" %*
  ) else (
    echo [错误] 未找到 Python。请先安装 Python 3.7 或更高版本：
    echo        https://www.python.org/downloads/
    echo        安装时请勾选 "Add Python to PATH"。
    pause
    exit /b 1
  )
)

REM 出错时不要把窗口直接关掉，方便看报错信息
if %errorlevel% neq 0 pause
