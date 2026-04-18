#!/bin/bash
# 定增项目企业基本面风险分析报告生成脚本
# 使用方法：./gen_report.sh <股票代码> [企业名称] [输出文件名]

set -e

# 检查参数
if [ $# -lt 1 ]; then
    echo "===================================="
    echo "定增项目企业基本面风险分析报告生成器"
    echo "===================================="
    echo ""
    echo "使用方法："
    echo "  $0 <股票代码> [企业名称] [输出文件名]"
    echo ""
    echo "示例："
    echo "  $0 002276.SZ                              # 企业名称自动获取"
    echo "  $0 603296.SH 华勤技术                      # 指定企业名称"
    echo "  $0 300735.SZ 光弘科技 自定义报告.docx        # 指定输出文件名"
    echo ""
    echo "股票代码格式："
    echo "  上交所：600xxx.SH 或 601xxx.SH"
    echo "  深交所：000xxx.SZ 或 002xxx.SZ 或 300xxx.SZ"
    echo "  北交所：8xxxxx.BJ"
    echo ""
    echo "说明："
    echo "  - 企业名称可选，不填则自动从Tushare获取"
    echo "  - 如不指定输出文件名，将自动生成：{企业名称}_定增项目企业基本面风险分析报告_{时间戳}.docx"
    echo ""
    exit 1
fi

# 获取参数
STOCK_CODE="$1"
STOCK_NAME="${2:-}"
OUTPUT_FILE="${3:-}"

# 生成时间戳
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")

# 输出目录（相对于脚本所在目录的 wanma_report/output/）
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
OUTPUT_DIR="${SCRIPT_DIR}/wanma_report/output"
mkdir -p "$OUTPUT_DIR"

# 如果未指定输出文件名，自动生成
if [ -z "$OUTPUT_FILE" ]; then
    if [ -n "$STOCK_NAME" ]; then
        OUTPUT_FILE="${STOCK_NAME}_定增项目企业基本面风险分析报告_${TIMESTAMP}.docx"
    else
        OUTPUT_FILE="${STOCK_CODE}_定增项目企业基本面风险分析报告_${TIMESTAMP}.docx"
    fi
fi

# 输出文件放到 output 目录
OUTPUT_PATH="${OUTPUT_DIR}/${OUTPUT_FILE}"

# 切换到脚本目录
cd "$(dirname "$0")/wanma_report"

echo "===================================="
echo "定增项目企业基本面风险分析报告生成器"
echo "===================================="
echo ""
echo "股票代码：$STOCK_CODE"
echo "企业名称：${STOCK_NAME:-自动获取}"
echo "输出文件：$OUTPUT_PATH"
echo ""

# 检查Tushare Token
if [ -z "$TUSHARE_TOKEN" ]; then
    echo "警告：未设置 TUSHARE_TOKEN 环境变量"
    echo "请在 config.yaml 中配置 token 或设置环境变量："
    echo "  export TUSHARE_TOKEN=your_token_here"
    echo ""
fi

# 构建命令参数
CMD_ARGS="--stock $STOCK_CODE"
if [ -n "$STOCK_NAME" ]; then
    CMD_ARGS="$CMD_ARGS --name $STOCK_NAME"
fi
CMD_ARGS="$CMD_ARGS --output $OUTPUT_PATH"

echo "开始生成报告..."
echo ""

# 运行报告生成
python3 main.py $CMD_ARGS

# 检查是否成功
if [ $? -eq 0 ]; then
    echo ""
    echo "===================================="
    echo "报告生成成功！"
    echo "输出文件：$OUTPUT_PATH"
    echo ""
    echo "报告包含以下内容："
    echo "   - 项目概况"
    echo "   - 风险概述"
    echo "   - 财务风险分析（四维度评分）"
    echo "   - 估值风险分析（DCF三种方法）"
    echo "   - 退出风险分析"
    echo "   - 招引落地建议"
    echo "===================================="
else
    echo ""
    echo "===================================="
    echo "报告生成失败"
    echo "请检查错误信息并重试"
    echo "===================================="
    exit 1
fi
