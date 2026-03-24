#!/bin/bash
# 测试所有章节模块

echo "========================================"
echo "🧪 模块化报告生成器 - 完整测试"
echo "========================================"

# 激活vnpy环境
source ~/anaconda3/etc/profile.d/conda.sh
conda activate vnpy

cd /Users/davy/github/investment_valuation/price_maintenance_risk_analysis/scripts/generate_word_report_v2

echo ""
echo "📋 测试配置："
echo "  股票代码: 300735.SZ"
echo "  股票名称: 光弘科技"
echo "  环境: vnpy"
echo ""

# 记录开始时间
START_TIME=$(date +%s)

# 运行报告生成
python3 main.py --stock 300735.SZ --name 光弘科技

# 记录结束时间
END_TIME=$(date +%s)
ELAPSED=$((END_TIME - START_TIME))

echo ""
echo "========================================"
echo "📊 测试结果"
echo "========================================"
echo "耗时: $ELAPSED 秒"

# 检查是否生成了报告文件
LATEST_REPORT=$(ls -t /Users/davy/github/investment_valuation/price_maintenance_risk_analysis/reports/outputs/*.docx 2>/dev/null | head -1)
if [ -n "$LATEST_REPORT" ]; then
    echo "✅ 报告已生成: $LATEST_REPORT"
    FILE_SIZE=$(du -h "$LATEST_REPORT" | cut -f1)
    echo "   文件大小: $FILE_SIZE"
else
    echo "❌ 报告生成失败"
fi

echo ""

# 退出vnpy环境
conda deactivate
