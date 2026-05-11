#!/bin/bash
# 通用定增风险分析报告生成脚本
# 使用方法：./gen_report.sh <股票代码> <企业名称> [报价日] [输出文件名]

set -e

# 检查参数
if [ $# -lt 2 ]; then
    echo "===================================="
    echo "定增风险分析报告生成器"
    echo "===================================="
    echo ""
    echo "使用方法："
    echo "  $0 <股票代码> <企业名称> [报价日] [输出文件名]"
    echo ""
    echo "示例："
    echo "  $0 603296.SH 华勤技术                      # 生成华勤技术报告（当前日期）"
    echo "  $0 300735.SZ 光弘科技 20260407            # 生成光弘科技报告（指定报价日）"
    echo "  $0 603296.SH 华勤技术 20260415 自定义.docx # 指定报价日和输出文件名"
    echo ""
    echo "股票代码格式："
    echo "  上交所：600xxx.SH 或 601xxx.SH"
    echo "  深交所：300xxx.SZ"
    echo "  北交所：8xxxxx.BJ"
    echo ""
    echo "说明："
    echo "  - 企业名称建议使用2-8个汉字，避免特殊字符"
    echo "  - 报价日格式：YYYYMMDD（如20260407），可选，默认使用当前日期"
    echo "  - 如不指定输出文件名，将自动生成：{企业名称}_定增市场风险分析报告_{时间戳}.docx"
    echo "  - 数据通过 SQLite 数据库管理，无需手动创建配置文件"
    echo ""
    exit 1
fi

# 获取参数
STOCK_CODE="$1"
STOCK_NAME="$2"
ISSUE_DATE="${3:-}"
OUTPUT_FILE="${4:-}"

# 判断第三个参数是报价日还是输出文件名
if [ -n "$ISSUE_DATE" ]; then
    # 检查是否为YYYYMMDD格式（8位数字）
    if echo "$ISSUE_DATE" | grep -qE '^[0-9]{8}$'; then
        # 是报价日
        ISSUE_DATE_PARAM="$ISSUE_DATE"
    else
        # 不是报价日格式，当作输出文件名处理
        OUTPUT_FILE="$ISSUE_DATE"
        ISSUE_DATE_PARAM=""
        ISSUE_DATE=""
    fi
else
    ISSUE_DATE_PARAM=""
fi

# 生成时间戳
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")

# 如果未指定输出文件名，自动生成带时间戳的文件名
if [ -z "$OUTPUT_FILE" ]; then
    OUTPUT_FILE="${STOCK_NAME}_定增市场风险分析报告_${TIMESTAMP}.docx"
fi

# 切换到脚本目录
cd "$(dirname "$0")"

echo "===================================="
echo "定增风险分析报告生成器"
echo "===================================="
echo ""
echo "股票代码：$STOCK_CODE"
echo "企业名称：$STOCK_NAME"
echo "输出文件：$OUTPUT_FILE"
echo ""
echo "当前配置："
echo "   投资金额: 1 亿元（固定，用于风险评估）"
if [ -n "$ISSUE_DATE" ]; then
    echo "   报价日: $ISSUE_DATE"
else
    echo "   报价日: 当前日期"
fi
echo ""
echo "开始生成报告..."
echo ""

# 运行报告生成脚本（数据通过DB管理，自动创建缺失配置）
cd scripts/generate_word_report_v2

if [ -n "$ISSUE_DATE_PARAM" ]; then
    # 有指定报价日
    python3 main.py \
        --stock "$STOCK_CODE" \
        --name "$STOCK_NAME" \
        --issue-date "$ISSUE_DATE_PARAM" \
        --output "$OUTPUT_FILE" \
        --force
else
    # 使用当前日期
    python3 main.py \
        --stock "$STOCK_CODE" \
        --name "$STOCK_NAME" \
        --output "$OUTPUT_FILE" \
        --force
fi

cd ../..

# 检查是否成功
if [ $? -eq 0 ]; then
    echo ""
    echo "===================================="
    echo "报告生成成功！"
    echo "输出文件：$OUTPUT_FILE"
    echo ""
    echo "报告包含以下内容："
    echo "   - 项目概况"
    echo "   - 相对估值分析"
    echo "   - DCF估值分析"
    echo "   - 敏感性分析"
    echo "   - 蒙特卡洛模拟"
    echo "   - 情景分析"
    echo "   - 压力测试"
    echo "   - VaR风险度量"
    echo "   - 综合评估"
    echo "   - 风控建议与风险提示"
    echo "===================================="
else
    echo ""
    echo "===================================="
    echo "报告生成失败"
    echo "请检查错误信息并重试"
    echo "===================================="
    exit 1
fi
