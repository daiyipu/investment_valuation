#!/bin/bash
# Git自动打标签脚本
# 使用方法: ./tag.sh [small|medium|big]

# 获取最新标签
latest_tag=$(git tag --list | sort -V | tail -1)

# 如果没有标签，从v1.0.0开始
if [ -z "$latest_tag" ]; then
    new_tag="v1.0.0"
    major=1; minor=0; patch=0
    echo "未找到标签，从v1.0.0开始"
else
    # 解析版本号（移除v前缀）
    version=${latest_tag#v}
    major=$(echo $version | cut -d. -f1)
    minor=$(echo $version | cut -d. -f2)
    patch=$(echo $version | cut -d. -f3)
    
    echo "当前标签: $latest_tag"
    
    # 判断改动大小
    change_type=${1:-"small"}
    
    case $change_type in
        big|major|大)
            echo "改动类型: 大改动 (+1.0.0)"
            major=$((major + 1))
            minor=0
            patch=0
            ;;
        medium|minor|中)
            echo "改动类型: 中改动 (+0.1.0)"
            minor=$((minor + 1))
            patch=0
            ;;
        small|patch|小|"")
            echo "改动类型: 小改动 (+0.0.1)"
            patch=$((patch + 1))
            ;;
        *)
            echo "错误: 无效的改动类型 '$change_type'"
            echo "使用方法: $0 [small|medium|big]"
            exit 1
            ;;
    esac
    
    new_tag="v${major}.${minor}.${patch}"
fi

# 删除刚才创建的错误标签
if [ "$latest_tag" != "" ] && [ "$new_tag" == "v1.0.1" ]; then
    git tag -d v1.0.1 2>/dev/null
fi

# 创建标签
git tag -a "$new_tag" -m "版本更新: ${latest_tag:-无} -> $new_tag"

echo "✅ 标签已创建: $new_tag"

# 询问是否推送
read -p "是否推送到远程? (y/n) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    git push origin "$new_tag"
    echo "✅ 标签已推送到远程"
else
    echo "⚠️  标签未推送，请手动执行: git push origin $new_tag"
fi
