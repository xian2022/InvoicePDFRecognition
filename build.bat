@echo off
chcp 65001 >nul
echo ========================================
echo   打包发票信息提取工具 (Windows版)
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
echo 安装依赖...
pip install -q -r requirements.txt
pip install -q pyinstaller

:: 打包
echo.
echo 开始打包...
pyinstaller --onefile ^
    --name "发票提取工具" ^
    --hidden-import=pdfplumber ^
    --hidden-import=openpyxl ^
    --hidden-import=flask ^
    app.py

echo.
if exist "dist\发票提取工具.exe" (
    echo ✅ 打包成功！
    echo 可执行文件位置: %CD%\dist\发票提取工具.exe
    echo.
    echo 使用方法:
    echo   双击 dist\发票提取工具.exe
    echo.
    echo 浏览器会自动打开 http://localhost:5000
) else (
    echo ❌ 打包失败，请检查错误信息
)

pause
