#!/bin/bash
# ============================================================
# 北京邮电大学学位论文一键格式化脚本
# 用法: ./format.sh 论文.docx
# 输出: 论文.docx (已格式化), 论文_backup.docx (原始备份)
# ============================================================

set -e

if [ $# -eq 0 ]; then
    echo "用法: ./format.sh <论文.docx>"
    echo "示例: ./format.sh 大论文框架.docx"
    exit 1
fi

INPUT="$1"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORK_DIR=$(mktemp -d)
BACKUP="${INPUT%.docx}_backup.docx"

if [ ! -f "$INPUT" ]; then
    echo "错误: 文件 '$INPUT' 不存在"
    exit 1
fi

echo "============================================================"
echo "  北京邮电大学学位论文格式化工具"
echo "============================================================"

# 1. 备份
cp "$INPUT" "$BACKUP"
echo "[1/5] 已备份原文件 → $BACKUP"

# 2. 解包
UNPACK_SCRIPT="$SCRIPT_DIR/.claude/skills/docx/scripts/office/unpack.py"
if [ -f "$UNPACK_SCRIPT" ]; then
    python3 "$UNPACK_SCRIPT" "$INPUT" "$WORK_DIR/" 2>&1 | sed 's/^/      /'
else
    mkdir -p "$WORK_DIR"
    unzip -o -q "$INPUT" -d "$WORK_DIR"
    echo "      (使用 unzip 解包)"
fi
echo "[2/5] 解包完成"

# 3. 格式化正文
python3 "$SCRIPT_DIR/format_thesis.py" "$WORK_DIR/word/document.xml" "$SCRIPT_DIR"
echo "[3/5] 正文格式化完成"

# 4. 格式化页眉页脚
python3 "$SCRIPT_DIR/format_headers_footers.py" "$WORK_DIR"
echo "[4/5] 页眉页脚格式化完成"

# 5. 重新打包
(cd "$WORK_DIR" && zip -r -q "$SCRIPT_DIR/$INPUT" . -x ".*")
echo "[5/5] 打包完成 → $INPUT"

# 清理
rm -rf "$WORK_DIR"

echo ""
echo "完成！请用 Word 打开 $INPUT 查看效果。"
echo "如需恢复，原文件在 $BACKUP"
