@echo off
chcp 65001 >nul
echo ========================================
echo   发票信息提取工具
echo ========================================

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

:: 进入脚本目录
cd /d "%~dp0"

:: 检查虚拟环境
if not exist "venv" (
    echo 正在创建虚拟环境...
    python -m venv venv
)

:: 激活虚拟环境
call venv\Scripts\activate.bat

:: 安装依赖
echo 检查并安装依赖...
pip install -q -r requirements.txt

:: 启动应用
echo.
echo 启动服务...
echo 请在浏览器中打开: http://localhost:5000
echo 按 Ctrl+C 停止服务
echo.

:: 延迟 2 秒后自动打开浏览器
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:5000"

python app.py
pause
