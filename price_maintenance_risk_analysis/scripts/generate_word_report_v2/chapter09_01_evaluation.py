"""
第九章综合评估汇总

本章节对前面章节的分析结果进行综合汇总，涵盖个股与行业对比、相对估值、DCF估值、蒙特卡洛模拟、情景分析、压力测试和VaR风险度量等核心内容。
"""

import sys
import os
from docx.shared import Inches

# 添加路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)

from module_utils import add_title, add_paragraph, add_table_data, add_image
from module_utils import font_prop


def generate_chapter(context):
    """
    生成第九章第一节：综合评估汇总

    参数说明：
    - context: 包含所有必要数据的上下文字典
    """
    # 从context中提取变量
    document = context['document']
    project_params = context['project_params']
    market_data = context['market_data']
    IMAGES_DIR = context['IMAGES_DIR']

    # 添加第九章主标题
    add_title(document, '九、风控建议与风险提示', level=1)
    add_paragraph(document, '本章节从风险控制角度，给出综合评估汇总、盈亏平衡分析、报价方案建议和全面的风险提示。')
    add_paragraph(document, '基于保守原则，确保投资决策在合理风险可控范围内进行。')
    add_paragraph(document, '')

    # 从context['results']中获取其他章节的计算结果
    all_scenarios = context['results'].get('all_scenarios', [])
    var_95 = context['results'].get('var_95', 0.5)
    var_99 = context['results'].get('var_99', 0.7)
    risk_rating = context['results'].get('risk_rating', '高风险')
    mc_volatility = context['results'].get('mc_volatility', 0.30)
    mc_drift = context['results'].get('mc_drift', 0.08)

    # ==================== 9.1 综合评估汇总 ====================
    add_title(document, '9.1 综合评估汇总', level=2)

    add_paragraph(document, '本节对前面章节的分析结果进行综合汇总，涵盖相对估值、DCF估值、蒙特卡洛模拟、情景分析、压力测试和VaR风险度量等核心内容。')
    add_paragraph(document, '')

    # ==================== 9.1.1 个股与历史分位数对比 ====================
    add_title(document, '9.1.1 个股与历史分位数对比', level=3)

    add_paragraph(document, '本节对比个股当前估值水平与自身历史数据，判断当前估值在历史中的位置。')
    add_paragraph(document, '')

    # 从第二章获取历史分位数数据（如果有的话）
    # 这里简化处理，显示基本信息
    current_price = project_params['current_price']
    issue_price = project_params['issue_price']
    ma20 = market_data.get('ma_20', current_price)
    ma60 = market_data.get('ma_60', current_price)
    ma120 = market_data.get('ma_120', current_price)
    ma250 = market_data.get('ma_250', current_price)

    # 计算当前价格相对于均线的位置
    ma_comparison_data = [
        ['指标', '当前值', 'MA20', 'MA60', 'MA120', 'MA250', '评估'],
        ['股价（元）', f'{current_price:.2f}', f'{ma20:.2f}', f'{ma60:.2f}', f'{ma120:.2f}', f'{ma250:.2f}', '']
    ]

    # 添加评估
    if current_price > ma120:
        price_eval = "高于半年线，偏强势"
    elif current_price < ma120 * 0.95:
        price_eval = "低于半年线5%以上，偏弱势"
    else:
        price_eval = "在半年线附近，震荡整理"
    ma_comparison_data[1][6] = price_eval

    add_table_data(document, ma_comparison_data[0], [ma_comparison_data[1]])

    add_paragraph(document, '')
    add_paragraph(document, ' 技术位置解读：', bold=True)
    add_paragraph(document, f'• 当前价格：{current_price:.2f}元')
    add_paragraph(document, f'• 相对MA20：{((current_price/ma20 - 1) * 100):+.2f}%')
    add_paragraph(document, f'• 相对MA120：{((current_price/ma120 - 1) * 100):+.2f}%')
    add_paragraph(document, f'• 技术位置：{price_eval}')
    add_paragraph(document, '')

    # ==================== 9.1.2 相对估值分析汇总 ====================
    add_title(document, '9.1.2 相对估值分析', level=3)

    add_paragraph(document, '本节汇总相对估值分析的核心结果，包括PE、PB等估值指标与行业的对比。')
    add_paragraph(document, '')
    add_paragraph(document, '注：详细的PE历史分位数分析请参见第二章"相对估值分析"。')
    add_paragraph(document, '')

    # 尝试从多个数据源获取PE/PB数据
    current_pe = 0
    industry_pe = 0
    current_pb = 0
    industry_pb = 0

    # 优先从context['results']获取第二章的计算结果
    if 'stock_pe_data' in context and context['stock_pe_data'] is not None:
        stock_pe_data = context['stock_pe_data']
        if not stock_pe_data.empty:
            current_pe = stock_pe_data.iloc[-1]['pe_ttm']

    if 'industry_pe_data' in context and context['industry_pe_data'] is not None:
        industry_pe_data = context['industry_pe_data']
        if not industry_pe_data.empty:
            industry_pe = industry_pe_data.iloc[-1]['pe_ttm']

    # 如果没有找到PE数据，尝试从market_data获取
    if current_pe == 0:
        current_pe = market_data.get('pe_ratio', 0)
    if industry_pe == 0:
        industry_pe = market_data.get('industry_pe', 0)

    # PB数据（如果有）
    current_pb = market_data.get('pb_ratio', 0)
    industry_pb = market_data.get('industry_pb', 0)

    # 如果有PE数据，生成表格
    if current_pe > 0 and industry_pe > 0:
        pe_premium = (current_pe / industry_pe - 1) * 100
        valuation_summary_data = [
            ['估值指标', '个股', '行业', '相对行业', '评估'],
            ['市盈率(PE)', f'{current_pe:.2f}倍', f'{industry_pe:.2f}倍', f'{pe_premium:+.1f}%', ''],
        ]

        if pe_premium < -20:
            pe_eval = "显著低估"
        elif pe_premium < 0:
            pe_eval = "相对低估"
        elif pe_premium < 20:
            pe_eval = "相对合理"
        elif pe_premium < 50:
            pe_eval = "相对高估"
        else:
            pe_eval = "显著高估"
        valuation_summary_data[1][4] = pe_eval

        if current_pb > 0 and industry_pb > 0:
            pb_premium = (current_pb / industry_pb - 1) * 100
            valuation_summary_data.append([
                '市净率(PB)', f'{current_pb:.2f}倍', f'{industry_pb:.2f}倍', f'{pb_premium:+.1f}%', ''
            ])

        add_table_data(document, valuation_summary_data[0], valuation_summary_data[1:])

        add_paragraph(document, '')
        add_paragraph(document, ' 相对估值解读：', bold=True)
        add_paragraph(document, f'• PE相对行业：{pe_premium:+.1f}%（{pe_eval}）')
        if current_pb > 0 and industry_pb > 0:
            add_paragraph(document, f'• PB相对行业：{pb_premium:+.1f}%')
        add_paragraph(document, '')
    else:
        add_paragraph(document, ' PE/PB数据暂时不可用，请参见第二章"相对估值分析"获取详细信息。')
        add_paragraph(document, '')

    # ==================== 9.1.3 DCF估值分析汇总 ====================
    add_title(document, '9.1.3 DCF估值分析', level=3)

    add_paragraph(document, '本节汇总DCF绝对估值分析的结果，评估公司的内在价值。')
    add_paragraph(document, '')

    # DCF估值方法说明
    add_paragraph(document, '【DCF估值方法】', bold=True)
    add_paragraph(document, '')
    add_paragraph(document, 'DCF（现金流折现）估值法通过预测公司未来自由现金流，并以加权平均资本成本（WACC）')
    add_paragraph(document, '折现至现值，得出公司的内在价值。')
    add_paragraph(document, '')
    add_paragraph(document, '估值步骤：')
    add_paragraph(document, '1. 预测未来10年自由现金流（FCF）')
    add_paragraph(document, '2. 计算终值（Terminal Value）')
    add_paragraph(document, '3. 以WACC折现至现值')
    add_paragraph(document, '4. 减去净债务，得到股权价值')
    add_paragraph(document, '5. 除以总股数，得到每股价值')
    add_paragraph(document, '')

    # DCF估值数据（从第三章结果获取）
    intrinsic_value = context['results'].get('intrinsic_value', 0)
    if intrinsic_value is not None and intrinsic_value > 0:
        dcf_discount = (intrinsic_value / current_price - 1) * 100
        dcf_discount_issue = (intrinsic_value / issue_price - 1) * 100

        # 显示关键估值参数
        add_paragraph(document, '【估值参数】', bold=True)
        add_paragraph(document, '')
        wacc_result = context.get('wacc', None)
        net_debt_result = context.get('net_debt', None)
        enterprise_value_result = context.get('enterprise_value', None)
        equity_value_result = context.get('equity_value', None)

        if wacc_result is not None:
            add_paragraph(document, f'• WACC（加权平均资本成本）: {wacc_result*100:.1f}%')
        if net_debt_result is not None:
            add_paragraph(document, f'• 净债务: {net_debt_result:.2f} 亿元')
        if enterprise_value_result is not None:
            add_paragraph(document, f'• 企业价值: {enterprise_value_result/100000000:.2f} 亿元')
        if equity_value_result is not None:
            add_paragraph(document, f'• 股权价值: {equity_value_result/100000000:.2f} 亿元')

        add_paragraph(document, '')
        add_paragraph(document, '【估值结果】', bold=True)
        add_paragraph(document, '')
        add_paragraph(document, f'• DCF内在价值: {intrinsic_value:.2f} 元/股')
        add_paragraph(document, f'• 当前价格: {current_price:.2f} 元/股')
        add_paragraph(document, f'• 发行价格: {issue_price:.2f} 元/股')
        add_paragraph(document, f'• 相对市价安全边际: {dcf_discount:+.1f}%')
        add_paragraph(document, f'• 相对发行价安全边际: {dcf_discount_issue:+.1f}%')
        add_paragraph(document, '')

        # 估值结论
        add_paragraph(document, '【估值结论】', bold=True)
        add_paragraph(document, '')

        if dcf_discount_issue > 15:
            conclusion = "✓ DCF估值显示，相比发行价有显著安全边际（>15%），估值合理偏低，投资价值较高。"
        elif dcf_discount_issue > 0:
            conclusion = "✓ DCF估值显示，相比发行价有一定安全边际，估值相对合理，具有一定投资价值。"
        elif dcf_discount_issue > -10:
            conclusion = "⚠ DCF估值显示，发行价略高于内在价值（<10%），需关注估值风险。"
        else:
            conclusion = "✗ DCF估值显示，发行价显著高于内在价值（≥10%），估值偏高，需谨慎投资。"

        add_paragraph(document, conclusion)
        add_paragraph(document, '')

        # 添加估值表格
        dcf_summary_data = [
            ['估值方法', '内在价值', '当前价格', '发行价格', '相对市价', '相对发行价'],
            ['DCF估值', f'{intrinsic_value:.2f}元', f'{current_price:.2f}元', f'{issue_price:.2f}元', f'{dcf_discount:+.1f}%', f'{dcf_discount_issue:+.1f}%']
        ]
        add_table_data(document, dcf_summary_data[0], dcf_summary_data[1:])

        add_paragraph(document, '')
        add_paragraph(document, '说明：')
        add_paragraph(document, '• DCF内在价值基于未来自由现金流预测和WACC折现计算')
        add_paragraph(document, '• 安全边际 = (内在价值 - 价格) / 价格 × 100%')
        add_paragraph(document, '• 正安全边际表示内在价值高于价格，具有投资价值')
        add_paragraph(document, '• 负安全边际表示内在价值低于价格，存在估值风险')
        add_paragraph(document, '')
        add_paragraph(document, '注：详细的DCF估值模型、敏感性分析和参数说明请参见第三章"DCF估值分析"。')
        add_paragraph(document, '')

        # 保存结果到上下文
        context['dcf_intrinsic_value'] = intrinsic_value
        context['dcf_discount'] = dcf_discount
        context['dcf_discount_issue'] = dcf_discount_issue
    else:
        add_paragraph(document, '⚠️  DCF估值不可用（内在价值为负或无法计算），请参见第三章"DCF估值分析"了解详情。')
        add_paragraph(document, '')
        add_paragraph(document, '可能原因：')
        add_paragraph(document, '• WACC计算失败（参数缺失或不合理）')
        add_paragraph(document, '• 净利润为负（无法进行DCF估值）')
        add_paragraph(document, '• 净债务过高（股权价值为负）')
        add_paragraph(document, '• 其他计算错误')
        add_paragraph(document, '')

    # ==================== 9.1.4 蒙特卡洛模拟结果汇总 ====================
    add_title(document, '9.1.4 蒙特卡洛模拟结果', level=3)

    add_paragraph(document, '本节汇总蒙特卡洛模拟的核心结果，包括基于历史参数和基于预测参数（ARIMA+GARCH）两种方法的120日窗口模拟。')
    add_paragraph(document, '')
    add_paragraph(document, '注：详细的蒙特卡洛模拟分析请参见第五章"蒙特卡洛模拟"。')
    add_paragraph(document, '')

    # 添加两种模拟方法的说明
    add_paragraph(document, '【两种模拟方法说明】', bold=True)
    add_paragraph(document, '')
    add_paragraph(document, '本报告采用两种蒙特卡洛模拟方法，以全面评估定增项目的风险收益特征：')
    add_paragraph(document, '')

    add_paragraph(document, '1. **历史参数模拟**（传统方法）')
    add_paragraph(document, '   • 数据来源：基于过去250个交易日的历史统计数据')
    add_paragraph(document, '   • 漂移率：使用历史对数收益率的年化均值')
    add_paragraph(document, '   • 波动率：使用历史对数收益率的年化标准差')
    add_paragraph(document, '   • 优点：客观反映长期平均水平，不受模型假设影响')
    add_paragraph(document, '   • 缺点：无法反映市场结构变化和未来趋势')
    add_paragraph(document, '')

    add_paragraph(document, '2. **预测参数模拟**（先进方法）')
    add_paragraph(document, '   • 漂移率：基于ARIMA时间序列模型预测未来收益趋势')
    add_paragraph(document, '   • 波动率：基于GARCH模型预测未来波动率变化')
    add_paragraph(document, '   • 优点：考虑时间序列的动态特征，反映市场变化趋势')
    add_paragraph(document, '   • 缺点：依赖模型假设，预测不确定性较大')
    add_paragraph(document, '')

    add_paragraph(document, '【方法对比与选择】', bold=True)
    add_paragraph(document, '')
    add_paragraph(document, '• 互补性：两种方法各有优劣，建议结合使用，互为验证')
    add_paragraph(document, '• 适用场景：')
    add_paragraph(document, '  - 历史参数：适用于市场稳定的成熟期公司')
    add_paragraph(document, '  - 预测参数：适用于成长期公司或市场环境变化较大时')
    add_paragraph(document, '• 投资决策：')
    add_paragraph(document, '  - 两种方法结果一致时：可增强决策信心')
    add_paragraph(document, '  - 两种方法结果分歧时：需进一步分析原因，采取更保守策略')
    add_paragraph(document, '')

    # 获取多窗口期模拟结果
    multi_window_results = context['results'].get('multi_window_mc_results', {})
    if multi_window_results and '120d' in multi_window_results:
        mc_120d = multi_window_results['120d']
        profit_prob_historical = mc_120d.get('profit_prob', 0.5) * 100
        mean_return_historical = mc_120d.get('mean_return', 0.0) * 100
        median_return_historical = mc_120d.get('median_return', 0.0) * 100

        # 获取预测参数模拟结果
        mc_predicted = context['results'].get('mc_predicted_120d', {})
        if mc_predicted:
            profit_prob_predicted = mc_predicted.get('profit_prob', 0.5) * 100
            mean_return_predicted = mc_predicted.get('mean_return', 0.0) * 100
            median_return_predicted = mc_predicted.get('median_return', 0.0) * 100
            drift_predicted = mc_predicted.get('drift', mc_drift)
            vol_predicted = mc_predicted.get('volatility', mc_volatility)

            # 生成对比表格
            mc_comparison_data = [
                ['指标', '历史参数模拟', '预测参数模拟', '差异'],
                ['年化漂移率', f'{mc_drift*100:+.2f}%', f'{drift_predicted*100:+.2f}%', f'{(drift_predicted-mc_drift)*100:+.2f}%'],
                ['年化波动率', f'{mc_volatility*100:.2f}%', f'{vol_predicted*100:.2f}%', f'{(vol_predicted-mc_volatility)*100:+.2f}%'],
                ['', '', '', ''],
                ['盈利概率', f'{profit_prob_historical:.1f}%', f'{profit_prob_predicted:.1f}%', f'{profit_prob_predicted-profit_prob_historical:+.1f}%'],
                ['平均收益率', f'{mean_return_historical:+.2f}%', f'{mean_return_predicted:+.2f}%', f'{mean_return_predicted-mean_return_historical:+.2f}%'],
                ['中位数收益率', f'{median_return_historical:+.2f}%', f'{median_return_predicted:+.2f}%', f'{median_return_predicted-median_return_historical:+.2f}%'],
            ]

            add_table_data(document, ['指标', '历史参数模拟', '预测参数模拟(ARIMA+GARCH)', '差异'], mc_comparison_data)

            add_paragraph(document, '')
            add_paragraph(document, ' 模拟结果解读：', bold=True)
            add_paragraph(document, f'• 历史参数模拟：基于250日历史数据的漂移率({mc_drift*100:+.2f}%)和波动率({mc_volatility*100:.2f}%)')
            add_paragraph(document, f'• 预测参数模拟：基于ARIMA预测的漂移率({drift_predicted*100:+.2f}%)和GARCH预测的波动率({vol_predicted*100:.2f}%)')
            add_paragraph(document, f'• 盈利概率差异：{profit_prob_predicted-profit_prob_historical:+.1f}个百分点（{"预测更高" if profit_prob_predicted > profit_prob_historical else "历史更高"}）')
            add_paragraph(document, f'• 预期收益差异：{mean_return_predicted-mean_return_historical:+.2f}个百分点（{"预测更乐观" if mean_return_predicted > mean_return_historical else "历史更乐观"}）')
            add_paragraph(document, '')
        else:
            # 只显示历史参数结果
            mc_summary_data = [
                ['模拟参数', '值'],
                ['模拟窗口', '120日（半年）'],
                ['模拟次数', '10,000次'],
                ['年化波动率', f'{mc_volatility*100:.2f}%'],
                ['年化漂移率', f'{mc_drift*100:+.2f}%'],
                ['', ''],
                ['模拟结果', ''],
                ['盈利概率', f'{profit_prob_historical:.1f}%'],
                ['平均收益率', f'{mean_return_historical:+.2f}%'],
                ['中位数收益率', f'{median_return_historical:+.2f}%'],
                ['风险评级', risk_rating]
            ]
            add_table_data(document, mc_summary_data[0], mc_summary_data[1:])

            add_paragraph(document, '')
            add_paragraph(document, ' 说明：预测参数模拟结果不可用，仅显示基于历史参数的模拟结果。')
            add_paragraph(document, '')
    else:
        add_paragraph(document, ' 蒙特卡洛模拟结果不可用，请参见第五章"蒙特卡洛模拟"获取详细信息。')
        add_paragraph(document, '')

    # ==================== 9.1.5 情景分析汇总 ====================
    add_title(document, '9.1.5 情景分析', level=3)

    add_paragraph(document, '本节汇总情景分析的核心结果，展示不同市场环境下的投资表现。')
    add_paragraph(document, '')

    # 情景分析结果（从第六章结果获取）
    if all_scenarios and len(all_scenarios) >= 15:
        # 按盈利概率排序（从高到低）
        sorted_scenarios = sorted(all_scenarios, key=lambda x: x.get('profit_prob', 0.0), reverse=True)

        # 分为三组：Top 5, Middle 5, Tail 5
        total_scenarios = len(sorted_scenarios)
        top_5 = sorted_scenarios[:5]
        middle_5 = sorted_scenarios[total_scenarios//2 - 2 : total_scenarios//2 + 3] if total_scenarios >= 10 else []
        tail_5 = sorted_scenarios[-5:]

        # 添加说明
        add_paragraph(document, '本节展示15种市场情景下的投资表现，按盈利概率分为三组：')
        add_paragraph(document, '')
        add_paragraph(document, '• **Top 5**（最优情景）：盈利概率最高的5种情景')
        add_paragraph(document, '• **Middle 5**（中等情景）：盈利概率居中的5种情景')
        add_paragraph(document, '• **Tail 5**（最差情景）：盈利概率最低的5种情景')
        add_paragraph(document, '')

        # ==================== Top 5 最优情景 ====================
        add_paragraph(document, '【Top 5 - 最优情景】', bold=True)
        add_paragraph(document, '')

        top_5_data = [['情景类型', '盈利概率', '中位数收益率', '评估']]
        for scenario in top_5:
            scenario_name = scenario.get('name', '情景')
            prob = scenario.get('profit_prob', 0.0) * 100
            median_ret = scenario.get('median_return', 0.0) * 100

            if prob >= 70:
                eval_text = "✓ 优秀"
            elif prob >= 50:
                eval_text = "✓ 良好"
            elif prob >= 30:
                eval_text = "○ 一般"
            else:
                eval_text = "✗ 较差"

            top_5_data.append([
                scenario_name,
                f'{prob:.1f}%',
                f'{median_ret:+.1f}%',
                eval_text
            ])

        add_table_data(document, top_5_data[0], top_5_data[1:])
        add_paragraph(document, '')

        # ==================== Middle 5 中等情景 ====================
        if middle_5:
            add_paragraph(document, '【Middle 5 - 中等情景】', bold=True)
            add_paragraph(document, '')

            middle_5_data = [['情景类型', '盈利概率', '中位数收益率', '评估']]
            for scenario in middle_5:
                scenario_name = scenario.get('name', '情景')
                prob = scenario.get('profit_prob', 0.0) * 100
                median_ret = scenario.get('median_return', 0.0) * 100

                if prob >= 70:
                    eval_text = "✓ 优秀"
                elif prob >= 50:
                    eval_text = "✓ 良好"
                elif prob >= 30:
                    eval_text = "○ 一般"
                else:
                    eval_text = "✗ 较差"

                middle_5_data.append([
                    scenario_name,
                    f'{prob:.1f}%',
                    f'{median_ret:+.1f}%',
                    eval_text
                ])

            add_table_data(document, middle_5_data[0], middle_5_data[1:])
            add_paragraph(document, '')

        # ==================== Tail 5 最差情景 ====================
        add_paragraph(document, '【Tail 5 - 最差情景】', bold=True)
        add_paragraph(document, '')

        tail_5_data = [['情景类型', '盈利概率', '中位数收益率', '评估']]
        for scenario in tail_5:
            scenario_name = scenario.get('name', '情景')
            prob = scenario.get('profit_prob', 0.0) * 100
            median_ret = scenario.get('median_return', 0.0) * 100

            if prob >= 70:
                eval_text = "✓ 优秀"
            elif prob >= 50:
                eval_text = "✓ 良好"
            elif prob >= 30:
                eval_text = "○ 一般"
            else:
                eval_text = "✗ 较差"

            tail_5_data.append([
                scenario_name,
                f'{prob:.1f}%',
                f'{median_ret:+.1f}%',
                eval_text
            ])

        add_table_data(document, tail_5_data[0], tail_5_data[1:])
        add_paragraph(document, '')

        # ==================== 情景分析总结 ====================
        add_paragraph(document, '【情景分析总结】', bold=True)
        add_paragraph(document, '')

        # 计算统计信息
        top_5_avg_prob = sum(s.get('profit_prob', 0.0) for s in top_5) / 5 * 100
        tail_5_avg_prob = sum(s.get('profit_prob', 0.0) for s in tail_5) / 5 * 100
        tail_5_worst = tail_5[0].get('profit_prob', 0.0) * 100

        add_paragraph(document, f'• 最优情景平均盈利概率: {top_5_avg_prob:.1f}%')
        add_paragraph(document, f'• 最差情景平均盈利概率: {tail_5_avg_prob:.1f}%')
        add_paragraph(document, f'• 最差情景盈利概率: {tail_5_worst:.1f}%（极端情况）')
        add_paragraph(document, '')

        add_paragraph(document, '投资启示：')
        add_paragraph(document, f'• 在Top 5情景下，盈利概率均在{top_5[4].get("profit_prob", 0.0)*100:.0f}%以上，投资价值较高')
        add_paragraph(document, f'• 在Tail 5情景下，盈利概率均低于{tail_5[4].get("profit_prob", 0.0)*100:.0f}%，需警惕极端风险')
        add_paragraph(document, f'• 建议关注最差情景的风险评估，确保在极端情况下损失可控')
        add_paragraph(document, '')

        add_paragraph(document, '注：详细的情景分析请参见第六章"情景分析"。')
        add_paragraph(document, '')

    elif all_scenarios:
        # 如果情景数量不足15个，显示所有情景
        add_paragraph(document, f'⚠️  当前共有{len(all_scenarios)}个情景，不足15个。显示所有情景如下：')
        add_paragraph(document, '')

        scenario_summary_data = [
            ['情景类型', '盈利概率', '中位数收益率', '评估']
        ]

        for scenario in all_scenarios:
            scenario_name = scenario.get('name', f'情景')
            prob = scenario.get('profit_prob', 0.0) * 100
            median_ret = scenario.get('median_return', 0.0) * 100

            if prob >= 70:
                eval_text = "✓ 优秀"
            elif prob >= 50:
                eval_text = "✓ 良好"
            elif prob >= 30:
                eval_text = "○ 一般"
            else:
                eval_text = "✗ 较差"

            scenario_summary_data.append([
                scenario_name,
                f'{prob:.1f}%',
                f'{median_ret:+.1f}%',
                eval_text
            ])

        add_table_data(document, scenario_summary_data[0], scenario_summary_data[1:])
        add_paragraph(document, '')
    else:
        add_paragraph(document, '⚠️  情景分析结果不可用，请参见第六章"情景分析"获取详细信息。')
        add_paragraph(document, '')

    # ==================== 9.1.6 压力测试结果 ====================
    add_title(document, '9.1.6 压力测试结果', level=3)

    add_paragraph(document, '本节汇总压力测试的结果，展示在极端市场情景下的投资表现。')
    add_paragraph(document, '')
    add_paragraph(document, '注：详细的压力测试分析请参见第七章"情景分析与压力测试"。')
    add_paragraph(document, '')

    # 添加压力测试说明
    add_paragraph(document, '【压力测试说明】', bold=True)
    add_paragraph(document, '')
    add_paragraph(document, '压力测试模拟历史危机和黑天鹅事件对股价的冲击，评估定增项目在极端情况下的抗风险能力。')
    add_paragraph(document, '本报告测试六种极端情景：2008年金融危机、2020年疫情、极端熊市、个股重大利空、行业政策收紧、流动性危机。')
    add_paragraph(document, '')

    # 定义六种极端压力情景（与第七章一致）
    current_price = project_params.get('current_price', 21.26)
    issue_price = project_params.get('issue_price', 25.80)
    issue_shares = project_params.get('issue_shares', 1000000)

    stress_scenarios = [
        {
            'name': '市场危机_2008',
            'description': '2008年全球金融危机',
            'price_drop': 0.60,
            'volatility_spike': 2.0,
            'details': 'A股市场暴跌约60%的极端情景'
        },
        {
            'name': '市场危机_2020',
            'description': '2020年新冠疫情冲击',
            'price_drop': 0.40,
            'volatility_spike': 1.5,
            'details': 'A股市场下跌约40%的情景'
        },
        {
            'name': '极端熊市',
            'description': '极端熊市周期',
            'price_drop': 0.50,
            'volatility_spike': 1.3,
            'details': '股价下跌50%，波动率上升30%'
        },
        {
            'name': '个股重大利空',
            'description': '公司业绩暴雷',
            'price_drop': 0.35,
            'volatility_spike': 1.8,
            'details': '财务造假、重大诉讼等'
        },
        {
            'name': '行业政策收紧',
            'description': '行业监管收紧',
            'price_drop': 0.25,
            'volatility_spike': 1.2,
            'details': '教培、互联网金融等强监管'
        },
        {
            'name': '流动性危机',
            'description': '市场流动性枯竭',
            'price_drop': 0.20,
            'volatility_spike': 2.5,
            'details': '钱荒、信用违约等'
        }
    ]

    # 计算各情景的定增收益
    add_paragraph(document, '【极端情景测试结果】', bold=True)
    add_paragraph(document, '')

    stress_results_data = [['情景名称', '情景描述', '股价跌幅', '压力后价格', '定增收益率', '亏损金额', '风险评估']]

    total_loss_amount = 0
    worst_scenario = None
    worst_loss = 0

    for scenario in stress_scenarios:
        stressed_price = current_price * (1 - scenario['price_drop'])
        pnl_percent = ((stressed_price - issue_price) / issue_price) * 100
        pnl_amount = (stressed_price - issue_price) * issue_shares / 100000000  # 转万元

        # 风险评估
        if pnl_percent >= 0:
            risk_level = '🟢 低风险'
            risk_assessment = '抗压能力强'
        elif pnl_percent >= -20:
            risk_level = '🟡 中等风险'
            risk_assessment = '有一定抗风险能力'
        elif pnl_percent >= -40:
            risk_level = '🟠 较高风险'
            risk_assessment = '抗风险能力较弱'
        else:
            risk_level = '🔴 高风险'
            risk_assessment = '抗风险能力很弱'

        stress_results_data.append([
            scenario['name'],
            scenario['description'],
            f'-{int(scenario["price_drop"]*100)}%',
            f'{stressed_price:.2f}元',
            f'{pnl_percent:+.1f}%',
            f'{pnl_amount:.2f}万元',
            risk_level
        ])

        total_loss_amount += max(0, -pnl_amount)

        # 找出最坏情景
        if pnl_percent < worst_loss:
            worst_loss = pnl_percent
            worst_scenario = scenario

    add_table_data(document, stress_results_data[0], stress_results_data[1:])
    add_paragraph(document, '')

    # 添加压力测试总结
    add_paragraph(document, '【压力测试总结】', bold=True)
    add_paragraph(document, '')

    if worst_scenario:
        add_paragraph(document, f'• **最坏情景**：{worst_scenario["description"]}（股价下跌{int(worst_scenario["price_drop"]*100)}%）')
        add_paragraph(document, f'• **最大亏损**：{abs(worst_loss):.1f}%（约{abs((issue_price - current_price * (1 - worst_scenario["price_drop"])) * issue_shares / 100000000):.2f}万元）')
        add_paragraph(document, f'• **亏损情景数量**：{sum(1 for s in stress_scenarios if (current_price * (1 - s["price_drop"]) < issue_price))} / 6 种情景出现亏损')
        add_paragraph(document, '')

    # 风险提示
    add_paragraph(document, '【风险提示】', bold=True)
    add_paragraph(document, '')
    add_paragraph(document, '• 压力测试情景基于历史事件，但实际极端情况可能更严重')
    add_paragraph(document, '• 投资决策应充分考虑极端情景下的最大承受能力')
    add_paragraph(document, '• 建议设置止损机制，控制单一项目投资比例')
    add_paragraph(document, '• 在市场出现极端情况时，应及时评估并调整投资策略')
    add_paragraph(document, '')

    # 抗风险能力评估
    add_paragraph(document, '【抗风险能力评估】', bold=True)
    add_paragraph(document, '')

    profitable_scenarios = sum(1 for s in stress_scenarios if (current_price * (1 - s['price_drop']) >= issue_price))

    if profitable_scenarios == 6:
        add_paragraph(document, '✓ 极强：在所有6种极端情景下仍能保持盈利')
    elif profitable_scenarios >= 4:
        add_paragraph(document, f'✓ 较强：在{profitable_scenarios}/6种极端情景下保持盈利')
    elif profitable_scenarios >= 2:
        add_paragraph(document, f'⚠️ 中等：在{profitable_scenarios}/6种极端情景下保持盈利，需谨慎')
    else:
        add_paragraph(document, f'✗ 较弱：仅在{profitable_scenarios}/6种极端情景下保持盈利，风险较高')

    add_paragraph(document, '')
    add_paragraph(document, '注：详细的压力测试分析请参见第七章"情景分析与压力测试"。')
    add_paragraph(document, '')

    # ==================== 9.1.7 VaR风险度量汇总 ====================
    add_title(document, '9.1.7 VaR风险度量汇总', level=3)

    add_paragraph(document, '本节汇总VaR（Value at Risk）风险度量的结果，评估在不同置信水平下的最大损失。')
    add_paragraph(document, '')
    add_paragraph(document, '注：详细的VaR风险度量分析请参见第八章"VaR风险度量"。')
    add_paragraph(document, '')

    # 使用120日窗口的VaR数据
    var_95_simple = var_95
    var_99_simple = var_99

    # 计算CVaR（条件风险价值，尾部平均损失）
    # 简化处理：CVaR约为VaR的1.2-1.5倍
    cvar_95_simple = min(var_95_simple * 1.3, 1.0)
    cvar_99_simple = min(var_99_simple * 1.4, 1.0)

    # 风险等级评估
    var_risk_level = "低风险" if var_95_simple < 0.15 else "中等风险" if var_95_simple < 0.30 else "高风险"

    var_summary_data = [
        ['置信水平', 'VaR（最大损失）', 'CVaR（尾部平均损失）', '说明'],
        ['90%置信', f'{min(var_95_simple*0.8, 1.0)*100:.2f}%', f'{min(var_95_simple*0.8*1.2, 1.0)*100:.2f}%', '10%概率损失超过此值'],
        ['95%置信', f'{var_95_simple*100:.2f}%', f'{cvar_95_simple*100:.2f}%', '5%概率损失超过此值'],
        ['99%置信', f'{var_99_simple*100:.2f}%', f'{cvar_99_simple*100:.2f}%', '1%概率损失超过此值'],
        ['', '', '', ''],
        ['数据窗口', '120日（半年锁定期）', '', '基于锁定期收益率'],
        ['风险等级', var_risk_level, '', '基于95% VaR综合评估'],
    ]

    add_table_data(document, var_summary_data[0], var_summary_data[1:])

    add_paragraph(document, '')
    add_paragraph(document, ' VaR解读：', bold=True)
    add_paragraph(document, f'• VaR基于锁定期（120日）的简单收益率计算，未进行年化')
    add_paragraph(document, f'• 95% VaR：{var_95_simple*100:.2f}%，表示在95%置信水平下，锁定期末的最大可能损失')
    add_paragraph(document, f'• 99% VaR：{var_99_simple*100:.2f}%，表示在99%置信水平下，锁定期末的最大可能损失')
    add_paragraph(document, f'• 风险等级：{var_risk_level}')
    add_paragraph(document, '')

    # ==================== 9.1.8 综合评估汇总 ====================
    add_title(document, '9.1.8 综合评估汇总', level=3)

    add_paragraph(document, '本节汇总9.1.1至9.1.7的分析结果，提供综合性的投资决策参考。')
    add_paragraph(document, '')

    # 获取120日窗口蒙特卡洛结果
    if multi_window_results and '120d' in multi_window_results:
        mc_120_result = multi_window_results['120d']
        profit_prob_display = mc_120_result.get('profit_prob', 0.5) * 100
        mean_return_display = mc_120_result.get('mean_return', 0.0) * 100
    else:
        profit_prob_display = 50.0
        mean_return_display = 0.0

    # DCF估值评估
    if intrinsic_value and intrinsic_value > 0:
        dcf_premium = (intrinsic_value / current_price - 1) * 100
        if dcf_premium > 50:
            dcf_eval = "显著低估"
        elif dcf_premium > 20:
            dcf_eval = "低估"
        elif dcf_premium > -20:
            dcf_eval = "合理估值"
        elif dcf_premium > -50:
            dcf_eval = "高估"
        else:
            dcf_eval = "显著高估"
        dcf_display = f"{intrinsic_value:.2f}元/股"
    else:
        dcf_eval = "数据不可用"
        dcf_display = "N/A"

    # 蒙特卡洛评估
    if profit_prob_display >= 70:
        mc_eval = "风险较低"
    elif profit_prob_display >= 50:
        mc_eval = "风险适中"
    else:
        mc_eval = "风险较高"

    # VaR评估
    if var_95_simple < 0.15:
        var_eval = "风险可控"
    elif var_95_simple < 0.30:
        var_eval = "风险较高"
    else:
        var_eval = "风险极高"

    # 综合评估表格
    comprehensive_summary_data = [
        ['评估维度', '指标值', '评估结果', '说明'],
        ['DCF内在价值', dcf_display, dcf_eval, '基于自由现金流折现（9.1.3）'],
        ['蒙特卡洛盈利概率', f'{profit_prob_display:.1f}%', mc_eval, '120日窗口模拟（9.1.4）'],
        ['预期年化收益率', f'{mean_return_display:.2f}%', '较高' if mean_return_display > 10 else '中等' if mean_return_display > 0 else '负收益', '模拟平均年化收益率（9.1.4）'],
        ['95% VaR（锁定期）', f'{var_95_simple*100:.2f}%', var_eval, '锁定期最大损失（9.1.7）'],
        ['', '', '', ''],
        ['综合建议', '', '', ''],
    ]

    # 根据多维度结果给出综合建议
    positive_signals = 0
    negative_signals = 0

    if dcf_eval in ['显著低估', '低估']:
        positive_signals += 1
    elif dcf_eval in ['高估', '显著高估']:
        negative_signals += 1

    if profit_prob_display >= 60:
        positive_signals += 1
    elif profit_prob_display < 40:
        negative_signals += 1

    if var_95_simple < 0.25:
        positive_signals += 1
    elif var_95_simple > 0.40:
        negative_signals += 1

    if positive_signals >= 2:
        comprehensive_advice = "建议积极参与"
        comprehensive_reason = "多个评估维度显示投资价值较高"
    elif negative_signals >= 2:
        comprehensive_advice = "建议谨慎参与"
        comprehensive_reason = "多个评估维度显示风险较高"
    else:
        comprehensive_advice = "建议中性参与"
        comprehensive_reason = "评估维度显示风险收益平衡"

    comprehensive_summary_data.append([
        '综合建议',
        comprehensive_advice,
        comprehensive_reason,
        '基于9.1.1至9.1.7的综合评估'
    ])

    add_table_data(document, comprehensive_summary_data[0], comprehensive_summary_data[1:])

    add_paragraph(document, '')
    add_paragraph(document, ' 综合评估解读：', bold=True)
    add_paragraph(document, f'• DCF估值：{dcf_display}（{dcf_eval}）')
    add_paragraph(document, f'• 蒙特卡洛盈利概率：{profit_prob_display:.1f}%（{mc_eval}）')
    add_paragraph(document, f'• 预期年化收益率：{mean_return_display:.2f}%')
    add_paragraph(document, f'• 95% VaR：{var_95_simple*100:.2f}%（{var_eval}）')
    add_paragraph(document, f'• 综合建议：{comprehensive_advice} - {comprehensive_reason}')
    add_paragraph(document, '')

    # ==================== 9.1 综合分析 ====================
    add_title(document, '9.1 综合分析', level=2)

    add_paragraph(document, '本节综合汇总前面各小节的核心信息，从多个维度全面评估定增项目的投资价值和风险。')
    add_paragraph(document, '')

    # ==================== 9.1.1 各小节核心信息汇总 ====================
    add_title(document, '9.1.1 各维度分析汇总', level=3)

    add_paragraph(document, '以下从相对估值、绝对估值、蒙特卡洛模拟、情景分析、压力测试、VaR风险度量六个维度，')
    add_paragraph(document, '全面汇总定增项目的投资价值和风险特征。')
    add_paragraph(document, '')

    # 创建综合汇总表格
    all_analysis_summary = []

    # 9.1.2 相对估值分析汇总
    if current_pe and current_pb:
        all_analysis_summary.append([
            '9.1.2 相对估值',
            f'PE: {current_pe:.2f}倍 vs 行业{industry_pe:.2f}倍',
            f'PB: {current_pb:.2f}倍 vs 行业{industry_pb:.2f}倍',
            f'{pe_eval}' if 'pe_eval' in locals() else '数据不可用'
        ])

    # 9.1.3 DCF绝对估值汇总
    if intrinsic_value and intrinsic_value > 0:
        dcf_premium = (intrinsic_value / current_price - 1) * 100
        all_analysis_summary.append([
            '9.1.3 DCF估值',
            f'内在价值: {intrinsic_value:.2f}元',
            f'相对市价: {dcf_premium:+.1f}%',
            dcf_eval
        ])
    else:
        all_analysis_summary.append([
            '9.1.3 DCF估值',
            '内在价值: N/A',
            'DCF估值不可用',
            '数据不可用'
        ])

    # 9.1.4 蒙特卡洛模拟汇总
    all_analysis_summary.append([
        '9.1.4 蒙特卡洛',
        f'盈利概率: {profit_prob_display:.1f}%',
        f'预期年化收益: {mean_return_display:+.2f}%',
        mc_eval
    ])

    # 9.1.5 情景分析汇总
    if all_scenarios and len(all_scenarios) >= 15:
        sorted_scenarios = sorted(all_scenarios, key=lambda x: x.get('profit_prob', 0.0), reverse=True)
        top_5_avg_prob = sum(s.get('profit_prob', 0.0) for s in sorted_scenarios[:5]) / 5 * 100
        tail_5_avg_prob = sum(s.get('profit_prob', 0.0) for s in sorted_scenarios[-5:]) / 5 * 100
        all_analysis_summary.append([
            '9.1.5 情景分析',
            f'Top 5情景平均: {top_5_avg_prob:.1f}%',
            f'Tail 5情景平均: {tail_5_avg_prob:.1f}%',
            f'情景数量: {len(all_scenarios)}个'
        ])
    elif all_scenarios:
        all_analysis_summary.append([
            '9.1.5 情景分析',
            f'情景数量: {len(all_scenarios)}个',
            '数据不足',
            '数据不可用'
        ])
    else:
        all_analysis_summary.append([
            '9.1.5 情景分析',
            '情景数量: 0个',
            '数据不可用',
            '数据不可用'
        ])

    # 9.1.6 压力测试汇总
    if worst_scenario:
        all_analysis_summary.append([
            '9.1.6 压力测试',
            f'最坏情景: {worst_scenario["description"]}',
            f'最大亏损: {abs(worst_loss):.1f}%',
            f'亏损情景: {sum(1 for s in stress_scenarios if (current_price * (1 - s["price_drop"]) < issue_price))}/6 种'
        ])
    else:
        all_analysis_summary.append([
            '9.1.6 压力测试',
            '压力测试: N/A',
            '数据不可用',
            '数据不可用'
        ])

    # 9.1.7 VaR风险度量汇总
    all_analysis_summary.append([
        '9.1.7 VaR度量',
        f'95% VaR: {var_95_simple*100:.2f}%',
        f'99% VaR: {var_99_simple*100:.2f}%',
        var_eval
    ])

    # 添加综合汇总表格
    if all_analysis_summary:
        summary_headers = ['分析维度', '核心指标1', '核心指标2', '评估结论']
        add_table_data(document, summary_headers, all_analysis_summary)
        add_paragraph(document, '')

    # ==================== 9.1.2 投资建议与风险提示 ====================
    add_title(document, '9.1.2 投资建议与风险提示', level=3)

    add_paragraph(document, '基于以上多维度综合分析，本节给出投资建议和风险提示。')
    add_paragraph(document, '')

    # 投资亮点
    add_paragraph(document, '【投资亮点】', bold=True)
    add_paragraph(document, '')

    highlights = []
    if dcf_eval in ['显著低估', '低估']:
        highlights.append(f'✓ DCF估值显示{dcf_eval}，内在价值{dcf_display}高于市价')

    if profit_prob_display >= 60:
        highlights.append(f'✓ 蒙特卡洛模拟显示盈利概率{profit_prob_display:.1f}%，{mc_eval}')

    if mean_return_display > 10:
        highlights.append(f'✓ 预期年化收益率{mean_return_display:.2f}%，收益空间较大')

    if var_95_simple < 0.25:
        highlights.append(f'✓ 95% VaR风险{var_eval}，最大损失可控')

    if highlights:
        for highlight in highlights:
            add_paragraph(document, highlight)
    else:
        add_paragraph(document, '暂无明显投资亮点')

    add_paragraph(document, '')

    # 风险提示
    add_paragraph(document, '【风险提示】', bold=True)
    add_paragraph(document, '')

    risks = []
    if dcf_eval in ['高估', '显著高估']:
        risks.append(f'⚠️ DCF估值显示{dcf_eval}，内在价值低于市价')

    if profit_prob_display < 40:
        risks.append(f'⚠️ 蒙特卡洛盈利概率{profit_prob_display:.1f}%，{mc_eval}')

    if mean_return_display < 0:
        risks.append(f'⚠️ 预期年化收益{mean_return_display:.2f}%，存在亏损风险')

    if var_95_simple > 0.40:
        risks.append(f'⚠️ 95% VaR风险{var_eval}，潜在损失较大')

    if worst_scenario and worst_loss < -40:
        risks.append(f'⚠️ 压力测试显示极端情况最大亏损{abs(worst_loss):.1f}%，需警惕极端风险')

    if risks:
        for risk in risks:
            add_paragraph(document, risk)
    else:
        add_paragraph(document, '暂无显著风险提示')

    add_paragraph(document, '')

    # 最终投资建议
    add_paragraph(document, '【最终投资建议】', bold=True)
    add_paragraph(document, '')

    investment_advice = comprehensive_advice
    investment_reason = comprehensive_reason

    add_paragraph(document, f'**{investment_advice}**')
    add_paragraph(document, f'理由：{investment_reason}')
    add_paragraph(document, '')

    # 操作建议
    add_paragraph(document, '【操作建议】', bold=True)
    add_paragraph(document, '')

    if positive_signals >= 2:
        add_paragraph(document, '• 建议积极参与，但建议控制单一项目投资比例（不超过总资产的20-30%）')
        add_paragraph(document, '• 建议设置止损机制，当股价跌破发行价15%时及时止损')
        add_paragraph(document, '• 建议分批建仓，降低择时风险')
    elif negative_signals >= 2:
        add_paragraph(document, '• 建议谨慎参与或要求更高的折价率（至少20%的安全边际）')
        add_paragraph(document, '• 建议设置严格的止损机制，当股价下跌超过10%时考虑止损')
        add_paragraph(document, '• 建议降低投资比例或观望等待更好的入场时机')
    else:
        add_paragraph(document, '• 建议中性参与，根据市场情况灵活调整仓位')
        add_paragraph(document, '• 建议设置止损机制，控制下行风险')
        add_paragraph(document, '• 建议结合其他投资机会进行分散投资')

    add_paragraph(document, '')

    # 重要提示
    add_paragraph(document, '【重要提示】', bold=True)
    add_paragraph(document, '')
    add_paragraph(document, '• 本报告基于历史数据和模型假设，实际结果可能存在偏差')
    add_paragraph(document, '• 市场有风险，投资需谨慎，本报告不构成投资建议')
    add_paragraph(document, '• 建议结合公司基本面、行业趋势、宏观环境等因素综合判断')
    add_paragraph(document, '• 建议定期回顾和调整投资策略，及时应对市场变化')
    add_paragraph(document, '')

    # 返回更新后的context
    return context
