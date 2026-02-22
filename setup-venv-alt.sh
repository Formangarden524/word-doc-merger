#!/bin/bash
#
# setup-venv-alt.sh - 备选方案：在用户主目录创建虚拟环境
#

set -e

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

VENV_DIR="$HOME/.word-doc-merger-venv"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo -e "${BLUE}🚀 在用户主目录创建虚拟环境...${NC}"
echo -e "${YELLOW}位置: $VENV_DIR${NC}"

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
echo "  source ~/.word-doc-merger-venv/bin/activate"
echo "  cd word-doc-merger"
echo "  python3 merge_docs_with_content.py <文件夹路径> <输出文件>"
