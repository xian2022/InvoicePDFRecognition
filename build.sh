#!/bin/bash
# 发票处理工具打包脚本

echo "========================================"
echo "  打包发票信息提取工具"
echo "========================================"

# 进入脚本目录
cd "$(dirname "$0")"

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 Python3"
    exit 1
fi

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo "正在创建虚拟环境..."
    python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
echo "安装依赖..."
pip install -q -r requirements.txt
pip install -q pyinstaller

# 打包（使用 --onefile 生成单文件，不包含 templates 目录）
echo ""
echo "开始打包..."
pyinstaller --onefile \
    --name "发票提取工具" \
    --hidden-import=pdfplumber \
    --hidden-import=openpyxl \
    --hidden-import=flask \
    app.py

echo ""
if [ -f "dist/发票提取工具" ]; then
    echo "✅ 打包成功！"
    echo "可执行文件位置: $(pwd)/dist/发票提取工具"
    echo ""
    echo "使用方法:"
    echo "  ./dist/发票提取工具"
    echo ""
    echo "然后在浏览器中打开 http://localhost:5000"
else
    echo "❌ 打包失败，请检查错误信息"
    exit 1
fi
