# 字体文件目录

## 用途

本目录用于存放中文字体文件，解决 matplotlib 中文显示乱码问题。

## 如何使用

### 方法 1: 手动下载（推荐）

1. 下载 SimHei.ttf（黑体）字体文件
   - GitHub 镜像: https://github.com/StellarCN/scp_zh/blob/master/fonts/SimHei.ttf
   - 或者从 Windows 系统: `C:\Windows\Fonts\simhei.ttf`

2. 将字体文件放到此目录
   ```bash
   cp /path/to/SimHei.ttf fonts/
   ```

3. 重启 Jupyter Notebook

### 方法 2: 使用下载脚本

```bash
python fonts/download_font.py
```

### 方法 3: 从系统复制

```bash
# macOS
cp /System/Library/Fonts/STHeiti\ Medium.ttc fonts/

# Linux (如果已安装)
cp /usr/share/fonts/truetype/wqy/wqy-microhei.ttc fonts/

# Windows
copy C:\Windows\Fonts\simhei.ttf fonts\
```

## 支持的字体

- `SimHei.ttf` - 黑体（推荐）
- `SimSun.ttf` - 宋体
- `msyh.ttc` - 微软雅黑
- 任何其他中文字体文件

## 注意事项

- 字体文件较大（几MB到十几MB），git 已配置忽略此目录
- 请勿将受版权保护的字体文件提交到 git 仓库
- 推荐使用开源字体或系统自带的字体
