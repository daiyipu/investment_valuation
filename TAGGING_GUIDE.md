# Git 标签管理指南

## 自动打标签功能

项目中已配置好自动打标签工具，根据改动大小自动递增版本号。

## 版本号规则

- **小改动 (patch)**: +0.0.1 (例如: v1.8.0 -> v1.8.1)
  - bug修复
  - 小功能改进
  - 文档更新
  
- **中改动 (minor)**: +0.1.0 (例如: v1.8.0 -> v1.9.0)
  - 新增功能
  - 模块添加
  - 重要改进
  
- **大改动 (major)**: +1.0.0 (例如: v1.8.0 -> v2.0.0)
  - 重大重构
  - 破坏性变更
  - 架构调整

## 使用方法

### 方法1: 提交后单独打标签

```bash
# 正常提交
git commit -m "fix: 修复中文乱码问题"

# 根据改动大小打标签
./tag.sh small      # 小改动 -> v1.8.1
./tag.sh medium     # 中改动 -> v1.9.0
./tag.sh big        # 大改动 -> v2.0.0
```

### 方法2: 使用Git别名

```bash
git commit -m "fix: 修复问题"
git tag-small       # 小改动
git tag-medium      # 中改动
git tag-big         # 大改动
```

### 方法3: 提交并打标签（一步完成）

```bash
./git-commit-tag.sh small "fix: 修复中文乱码问题"
./git-commit-tag.sh medium "feat: 添加新功能"
./git-commit-tag.sh big "refactor: 重构核心模块"
```

## 提交信息规范

建议使用 Conventional Commits 格式：

- `fix:` - bug修复（小改动）
- `feat:` - 新功能（中改动）
- `refactor:` - 重构（根据大小）
- `docs:` - 文档更新（小改动）
- `style:` - 代码格式（小改动）
- `perf:` - 性能优化（小改动）
- `test:` - 测试相关（小改动）
- `chore:` - 构建/工具（小改动）

## 当前版本

最新标签: v1.8.1

## 查看所有标签

```bash
git tag --list
git tag --list | sort -V  # 按版本号排序
```

## 删除标签（如需）

```bash
# 删除本地标签
git tag -d v1.8.1

# 删除远程标签
git push origin :refs/tags/v1.8.1
```
