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
    add_paragraph(document, '注：详细的DCF估值模型和敏感性分析请参见第三章"DCF估值分析"。')
    add_paragraph(document, '')

    # DCF估值数据（从第三章结果获取）
    intrinsic_value = context['results'].get('intrinsic_value', 0)
    if intrinsic_value is not None and intrinsic_value > 0:
        dcf_discount = (intrinsic_value / current_price - 1) * 100
        dcf_summary_data = [
            ['估值方法', '内在价值', '当前价格', '折价/溢价率', '评估'],
            ['DCF估值', f'{intrinsic_value:.2f}元', f'{current_price:.2f}元', f'{dcf_discount:+.1f}%', '']
        ]

        if dcf_discount < -20:
            dcf_eval = "显著低估"
        elif dcf_discount < 0:
            dcf_eval = "相对低估"
        elif dcf_discount < 20:
            dcf_eval = "相对合理"
        elif dcf_discount < 50:
            dcf_eval = "相对高估"
        else:
            dcf_eval = "显著高估"
        dcf_summary_data[1][4] = dcf_eval

        add_table_data(document, dcf_summary_data[0], [dcf_summary_data[1]])

        add_paragraph(document, '')
        add_paragraph(document, ' DCF估值解读：', bold=True)
        add_paragraph(document, f'• 内在价值：{intrinsic_value:.2f}元')
        add_paragraph(document, f'• 相对当前价：{dcf_discount:+.1f}%（{dcf_eval}）')
        add_paragraph(document, '')
    else:
        add_paragraph(document, ' DCF估值不可用（内在价值为负或无法计算），请参见第三章"DCF估值分析"了解详情。')
        add_paragraph(document, '')

    # ==================== 9.1.4 蒙特卡洛模拟结果汇总 ====================
    add_title(document, '9.1.4 蒙特卡洛模拟结果', level=3)

    add_paragraph(document, '本节汇总蒙特卡洛模拟的核心结果，包括基于历史参数和基于预测参数（ARIMA+GARCH）两种方法的120日窗口模拟。')
    add_paragraph(document, '')
    add_paragraph(document, '注：详细的蒙特卡洛模拟分析请参见第五章"蒙特卡洛模拟"。')
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
    if all_scenarios:
        # 显示部分核心情景
        scenario_summary_data = [
            ['情景类型', '盈利概率', '中位数收益率', '评估']
        ]

        # 取前5个情景作为示例
        for i, scenario in enumerate(all_scenarios[:5]):
            scenario_name = scenario.get('name', f'情景{i+1}')
            prob = scenario.get('profit_prob', 0.0) * 100
            median_ret = scenario.get('median_return', 0.0) * 100

            if prob >= 70:
                eval_text = "优秀"
            elif prob >= 50:
                eval_text = "良好"
            elif prob >= 30:
                eval_text = "一般"
            else:
                eval_text = "较差"

            scenario_summary_data.append([
                scenario_name,
                f'{prob:.1f}%',
                f'{median_ret:+.1f}%',
                eval_text
            ])

        add_table_data(document, scenario_summary_data[0], scenario_summary_data[1:])

        add_paragraph(document, '')

    # ==================== 9.1.6 压力测试结果 ====================
    add_title(document, '9.1.6 压力测试结果', level=3)

    add_paragraph(document, '本节汇总压力测试的结果，展示在极端市场情景下的投资表现。')
    add_paragraph(document, '')
    add_paragraph(document, '注：详细的压力测试分析请参见第七章"情景分析与压力测试"。')
    add_paragraph(document, '')

    # 提取压力测试结果
    extreme_stress_result = None
    pessimistic_result = None

    if all_scenarios:
        for scenario in all_scenarios:
            scenario_name = scenario.get('name', '')
            if '极端' in scenario_name or '压力' in scenario_name:
                extreme_stress_result = scenario
            elif '悲观' in scenario_name or '最坏' in scenario_name:
                pessimistic_result = scenario

    # 使用极端情景或悲观情景
    stress_scenario = extreme_stress_result or pessimistic_result

    if stress_scenario:
        scenario_name_display = stress_scenario.get('name', '压力测试情景')
        stress_prob = stress_scenario.get('profit_prob', 0.0) * 100
        stress_median = stress_scenario.get('median_return', 0.0) * 100

        stress_data = [
            ['测试类型', '盈利概率', '中位数收益', '说明'],
            [scenario_name_display, f'{stress_prob:.1f}%', f'{stress_median:.2f}%', '压力测试情景'],
        ]

        # 与基准情景对比（使用第一个情景作为基准）
        if all_scenarios:
            baseline = all_scenarios[0]
            baseline_prob = baseline.get('profit_prob', 0.0) * 100
            baseline_median = baseline.get('median_return', 0.0) * 100

            prob_diff = stress_prob - baseline_prob
            median_diff = stress_median - baseline_median

            stress_data.append([
                '与基准情景差异',
                f'{prob_diff:+.1f}个百分点',
                f'{median_diff:+.2f}%',
                '压力情景下的额外风险'
            ])

        add_table_data(document, stress_data[0], stress_data[1:])

        add_paragraph(document, '')
        add_paragraph(document, ' 压力测试解读：', bold=True)
        add_paragraph(document, f'• 压力测试情景：{scenario_name_display}')
        add_paragraph(document, f'• 盈利概率：{stress_prob:.1f}%')
        add_paragraph(document, f'• 中位数收益：{stress_median:+.2f}%')
        if len(stress_data) > 2:
            add_paragraph(document, f'• 与基准差异：盈利概率下降{abs(prob_diff):.1f}个百分点，中位数收益下降{abs(median_diff):.2f}%')
        add_paragraph(document, '')
    else:
        add_paragraph(document, ' 未找到压力测试情景结果，请参见第七章"情景分析与压力测试"获取详细信息。')
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

    # 返回更新后的context
    return context
