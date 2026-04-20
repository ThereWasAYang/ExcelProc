@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ==========================================
echo ExcelProc 一键环境配置
echo ==========================================
echo.

set "PYTHON_CMD="

py -3.12 -V >nul 2>nul && set "PYTHON_CMD=py -3.12"
if not defined PYTHON_CMD py -3 -V >nul 2>nul && set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD python -V >nul 2>nul && set "PYTHON_CMD=python"

if not defined PYTHON_CMD (
    echo [错误] 未检测到可用的 Python。
    echo.
    echo 请先安装 Python 3.12 或更高版本，然后重新运行本脚本。
    echo 下载地址：https://www.python.org/downloads/
    echo 安装时建议勾选 "Add python.exe to PATH"。
    echo.
    pause
    exit /b 1
)

echo [1/4] 使用以下命令创建虚拟环境：
echo     %PYTHON_CMD%
echo.

if not exist ".venv\Scripts\python.exe" (
    %PYTHON_CMD% -m venv ".venv"
    if errorlevel 1 (
        echo [错误] 创建虚拟环境失败。
        echo.
        pause
        exit /b 1
    )
) else (
    echo 检测到已有虚拟环境，跳过创建。
)

set "VENV_PYTHON=%CD%\.venv\Scripts\python.exe"

if not exist "%VENV_PYTHON%" (
    echo [错误] 虚拟环境中的 Python 不存在：%VENV_PYTHON%
    echo.
    pause
    exit /b 1
)

echo.
echo [2/4] 升级 pip ...
"%VENV_PYTHON%" -m pip install --upgrade pip
if errorlevel 1 (
    echo [错误] pip 升级失败。
    echo.
    pause
    exit /b 1
)

echo.
echo [3/4] 安装项目依赖 ...
"%VENV_PYTHON%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [错误] 依赖安装失败。
    echo.
    pause
    exit /b 1
)

echo.
echo [4/4] 创建必要目录 ...
if not exist "inputs" mkdir "inputs"
if not exist "configs" mkdir "configs"
if not exist "outputs" mkdir "outputs"

echo.
echo ==========================================
echo 环境配置完成
echo ==========================================
echo.
echo 虚拟环境位置：
echo   .venv
echo.
echo 以后可使用以下命令运行：
echo   .venv\Scripts\python.exe .\excel_processor.py --config .\configs\sample_config.jsonc
echo.
echo 如果你使用命令行，也可以先激活环境：
echo   .venv\Scripts\activate
echo.
pause
