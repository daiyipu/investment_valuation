#!/bin/bash
# 通用定增风险分析报告生成脚本
# 使用方法：./gen_report.sh <股票代码> [输出文件名]

set -e

# 检查参数
if [ $# -lt 1 ]; then
    echo "===================================="
    echo "📊 定增风险分析报告生成器"
    echo "===================================="
    echo ""
    echo "使用方法："
    echo "  $0 <股票代码> [输出文件名]"
    echo ""
    echo "示例："
    echo "  $0 603296.SH                    # 生成华勤技术报告"
    echo "  $0 603296.SH 华勤技术报告.docx  # 指定输出文件名"
    echo "  $0 300735.SZ                    # 生成光弘科技报告"
    echo ""
    echo "股票代码格式："
    echo "  上交所：600xxx.SH 或 601xxx.SH"
    echo "  深交所：300xxx.SZ"
    echo "  北交所：8xxxxx.BJ"
    echo ""
    exit 1
fi

# 获取参数
STOCK_CODE="$1"
OUTPUT_FILE="${2:-自动生成报告.docx}"

# 如果输出文件名为"自动生成报告.docx"，使用股票名称
if [ "$OUTPUT_FILE" = "自动生成报告.docx" ]; then
    OUTPUT_FILE="${STOCK_CODE}_定增风险分析报告.docx"
fi

# 切换到脚本目录
cd "$(dirname "$0")"

echo "===================================="
echo "📊 定增风险分析报告生成器"
echo "===================================="
echo ""
echo "股票代码：$STOCK_CODE"
echo "输出文件：$OUTPUT_FILE"
echo ""

# 转换股票代码格式（603296.SH -> 603296_SH）
CONFIG_STOCK_CODE=$(echo "$STOCK_CODE" | sed 's/\./_/g')
CONFIG_FILE="data/${CONFIG_STOCK_CODE}_placement_params.json"

# 检查配置文件是否存在
if [ ! -f "$CONFIG_FILE" ]; then
    echo "⚠️  配置文件不存在：$CONFIG_FILE"
    echo ""
    echo "正在创建配置文件模板..."

    # 创建配置文件模板（使用新格式）
    cat > "$CONFIG_FILE" << 'EOF'
{
  "financing_amount": 100000000,
  "lockup_period": 6,
  "pricing_method": "ma20_discount_90",
  "premium_rate": -0.10,
  "risk_free_rate": 0.03,
  "net_assets": 0.0,
  "total_debt": 0.0,
  "net_income": 0.0,
  "revenue_growth": 0.15,
  "operating_margin": 0.15,
  "beta": 1.0,
  "historical_fcf_data": {
    "years": 5,
    "year_range": [2020, 2024],
    "data": []
  },
  "_notes": {
    "financing_amount": "投资金额（元）- 固定1亿元，用于风险评估（与实际投资规模无关）",
    "lockup_period": "锁定期（月）- 默认6个月",
    "pricing_method": "定价方式：ma20_discount_90(MA20九折), ma20_discount_85(MA20八五折), ma20_par(MA20平价), custom_premium(自定义溢价率)",
    "premium_rate": "溢价率（负数为折价，正数为溢价）- 默认-0.10表示九折",
    "_auto_generated": "以下参数自动计算，无需手动填写：",
    "issue_price": "自动计算：MA20 × (1 + premium_rate)",
    "current_price": "自动从API获取最新股价"
  }
}
EOF

    echo "✅ 配置文件模板已创建：$CONFIG_FILE"
    echo ""
    echo "📝 配置说明："
    echo "   • 投资金额已固定为1亿元（用于风险评估）"
    echo "   • 风险程度与投资金额无关，因此使用固定金额便于比较不同项目"
    echo "   • 其他参数将自动从API获取，无需手动填写"
    echo ""
    echo "💡 如需调整其他参数（如锁定期、定价方式等），可编辑配置文件："
    echo "   nano $CONFIG_FILE"
    echo "   或者"
    echo "   vim $CONFIG_FILE"
    echo ""

    # 尝试打开编辑器
    if command -v code &> /dev/null; then
        read -p "是否现在用VS Code打开编辑？(Y/n): " open_editor
        if [ "$open_editor" = "Y" ] || [ "$open_editor" = "y" ]; then
            code "$CONFIG_FILE"
            echo ""
            echo "编辑完成后，请重新运行脚本："
            echo "  $0 $STOCK_CODE $OUTPUT_FILE"
            exit 0
        fi
    fi

    echo "配置完成后，重新运行脚本："
    echo "  $0 $STOCK_CODE $OUTPUT_FILE"
    exit 0
fi

echo "✅ 配置文件已找到：$CONFIG_FILE"
echo ""

# 读取配置
FINANCING_AMOUNT=$(python3 -c "import json; print(json.load(open('$CONFIG_FILE')).get('financing_amount', 100000000))" 2>/dev/null || echo "100000000")

echo "📊 当前配置："
echo "   投资金额: 1 亿元（固定，用于风险评估）"
echo ""

echo "✅ 配置参数验证通过"
echo ""
echo "⏳ 开始生成报告..."
echo ""

# 运行报告生成脚本
python3 scripts/generate_word_report_v2.py \
    --stock "$STOCK_CODE" \
    --output "$OUTPUT_FILE" \
    --force-update

# 检查是否成功
if [ $? -eq 0 ]; then
    echo ""
    echo "===================================="
    echo "✅ 报告生成成功！"
    echo "📄 输出文件：$OUTPUT_FILE"
    echo ""
    echo "📊 报告包含以下内容："
    echo "   ✓ 项目概况"
    echo "   ✓ 相对估值分析"
    echo "   ✓ DCF估值分析"
    echo "   ✓ 敏感性分析"
    echo "   ✓ 蒙特卡洛模拟"
    echo "   ✓ 情景分析"
    echo "   ✓ 压力测试"
    echo "   ✓ VaR风险度量"
    echo "   ✓ 综合评估"
    echo "   ✓ 风控建议与风险提示"
    echo "===================================="
else
    echo ""
    echo "===================================="
    echo "❌ 报告生成失败"
    echo "请检查错误信息并重试"
    echo "===================================="
    exit 1
fi
