#!/bin/bash
#
# setup-venv.sh - 创建隔离的虚拟环境运行 word-doc-merger
#

set -e

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/.venv"

echo -e "${BLUE}🚀 创建虚拟环境...${NC}"

# 创建虚拟环境
python3 -m venv "$VENV_DIR"

# 激活虚拟环境
source "$VENV_DIR/bin/activate"

# 安装依赖
echo -e "${YELLOW}📦 安装依赖...${NC}"
pip install python-docx

echo -e "${GREEN}✅ 虚拟环境创建完成！${NC}"
echo ""
echo -e "${BLUE}使用方法:${NC}"
echo "  source .venv/bin/activate"
echo "  ./merge-docs <文件夹路径>"
echo ""
echo "或直接使用（自动激活虚拟环境）:"
echo "  ./merge-docs-venv <文件夹路径>"
