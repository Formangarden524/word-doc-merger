#!/bin/bash
#
# install.sh - 安装 merge-docs 到 macOS 系统
#

set -e

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INSTALL_DIR="/usr/local/bin"

echo -e "${BLUE}🚀 Word Doc Merger 安装程序${NC}"
echo ""

# 检查是否已安装
if [ -f "$INSTALL_DIR/merge-docs" ]; then
    echo -e "${YELLOW}⚠️  merge-docs 已存在，将覆盖安装${NC}"
fi

# 创建安装目录（如果不存在）
if [ ! -d "$INSTALL_DIR" ]; then
    echo -e "${BLUE}📁 创建安装目录...${NC}"
    sudo mkdir -p "$INSTALL_DIR"
fi

# 复制脚本
echo -e "${BLUE}📦 安装 merge-docs...${NC}"
sudo cp "$SCRIPT_DIR/merge-docs" "$INSTALL_DIR/"
sudo chmod +x "$INSTALL_DIR/merge-docs"

# 复制 Python 脚本到同一目录
sudo cp "$SCRIPT_DIR/merge_docs_with_content.py" "$INSTALL_DIR/"

echo ""
echo -e "${GREEN}✅ 安装完成！${NC}"
echo ""
echo -e "${BLUE}使用方法:${NC}"
echo "  merge-docs <输入文件夹> [输出文件]"
echo ""
echo -e "${BLUE}示例:${NC}"
echo "  merge-docs ~/Documents/my_docs"
echo "  merge-docs ~/Downloads/docs ~/Desktop/output.docx"
echo ""
echo -e "${YELLOW}提示: 重新打开终端或运行 'source ~/.zshrc' 使命令生效${NC}"
