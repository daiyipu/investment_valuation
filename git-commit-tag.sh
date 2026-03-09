#!/bin/bash
# Git提交并打标签脚本
# 使用方法: ./git-commit-tag.sh [small|medium|big] "提交信息"

if [ $# -lt 2 ]; then
    echo "使用方法: $0 [small|medium|big] \"提交信息\""
    echo ""
    echo "示例:"
    echo "  $0 small \"fix: 修复中文乱码问题\""
    echo "  $0 medium \"feat: 添加新功能\""
    echo "  $0 big \"refactor: 重构核心模块\""
    exit 1
fi

change_type=$1
commit_msg=$2

# 执行git提交
git commit -m "$commit_msg"

if [ $? -ne 0 ]; then
    echo "❌ Git提交失败"
    exit 1
fi

# 执行打标签
./tag.sh $change_type

echo ""
echo "✅ 提交和打标签完成"
