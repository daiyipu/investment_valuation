#!/bin/bash
# 上市公司估值分析报告生成脚本
# 使用方法：./gen_report.sh <股票代码> [企业名称] [章节号]

set -e

# 检查参数
if [ $# -lt 1 ]; then
    echo "===================================="
    echo "  上市公司估值分析报告生成器"
    echo "===================================="
    echo ""
    echo "使用方法："
    echo "  $0 <股票代码> [企业名称] [章节号]"
    echo ""
    echo "示例："
    echo "  $0 002276.SZ                          # 企业名称自动获取"
    echo "  $0 601127 赛力斯                       # 自动补后缀.SH"
    echo "  $0 300750.SZ 宁德时代                   # 指定企业名称"
    echo "  $0 002276.SZ 万马高分子 4               # 仅生成第4章（调试用）"
    echo ""
    echo "股票代码格式（支持自动修正后缀）："
    echo "  上交所：600xxx.SH / 601xxx.SH / 603xxx.SH"
    echo "  深交所：000xxx.SZ / 002xxx.SZ / 300xxx.SZ"
    echo "  输入纯数字也可，6xx自动补.SH，0xx/3xx自动补.SZ"
    echo ""
    echo "报告包含11章："
    echo "  1.项目概述  2.财务分析  3.行业分析  4.DCF绝对估值"
    echo "  5.相对估值  6.敏感性分析  7.蒙特卡洛模拟  8.情景分析"
    echo "  9.压力测试  10.VaR风险度量  11.综合评估"
    echo ""
    exit 1
fi

# 获取参数
STOCK_CODE="$1"
STOCK_NAME="${2:-}"
CHAPTER="${3:-}"

# 切换到脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

PYTHON="/Users/davy/anaconda3/envs/vnpy/bin/python"

echo "===================================="
echo "  上市公司估值分析报告生成器"
echo "===================================="
echo ""
echo "股票代码：$STOCK_CODE"
echo "企业名称：${STOCK_NAME:-自动获取}"
if [ -n "$CHAPTER" ]; then
    echo "生成章节：第${CHAPTER}章（仅）"
else
    echo "生成章节：全部11章"
fi
echo ""

# 检查Tushare Token
if [ -z "$TUSHARE_TOKEN" ]; then
    if [ ! -f config.yaml ]; then
        echo "警告：未设置 TUSHARE_TOKEN 环境变量且未找到 config.yaml"
        echo "请设置环境变量：export TUSHARE_TOKEN=your_token_here"
        echo ""
    fi
fi

# 构建命令参数
CMD_ARGS="--stock $STOCK_CODE"
if [ -n "$STOCK_NAME" ]; then
    CMD_ARGS="$CMD_ARGS --name $STOCK_NAME"
fi
if [ -n "$CHAPTER" ]; then
    CMD_ARGS="$CMD_ARGS --chapter $CHAPTER"
fi

echo "开始生成报告..."
echo ""

# 运行报告生成
$PYTHON main.py $CMD_ARGS

# 检查是否成功
if [ $? -eq 0 ]; then
    echo ""
    echo "===================================="
    echo "  报告生成成功！"
    echo "  输出目录：${SCRIPT_DIR}/output/"
    echo ""
    echo "  报告内容："
    echo "   1. 项目概述      2. 财务分析      3. 行业分析"
    echo "   4. DCF绝对估值   5. 相对估值      6. 敏感性分析"
    echo "   7. 蒙特卡洛模拟  8. 情景分析      9. 压力测试"
    echo "  10. VaR风险度量  11. 综合评估"
    echo "===================================="
else
    echo ""
    echo "===================================="
    echo "  报告生成失败"
    echo "  请检查错误信息并重试"
    echo "===================================="
    exit 1
fi
