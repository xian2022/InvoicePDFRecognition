#!/bin/bash
# 发票处理工具启动脚本

echo "========================================"
echo "  发票信息提取工具"
echo "========================================"

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 Python3，请先安装 Python 3.8+"
    exit 1
fi

# 进入脚本目录
cd "$(dirname "$0")"

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo "正在创建虚拟环境..."
    python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
echo "检查并安装依赖..."
pip install -q -r requirements.txt

# 启动应用
echo ""
echo "启动服务..."
echo "请在浏览器中打开: http://localhost:5000"
echo "按 Ctrl+C 停止服务"
echo ""

# 延迟 2 秒后自动打开浏览器（后台执行）
(sleep 2 && open "http://localhost:5000" 2>/dev/null || xdg-open "http://localhost:5000" 2>/dev/null || echo "请手动打开浏览器访问: http://localhost:5000") &

python app.py
