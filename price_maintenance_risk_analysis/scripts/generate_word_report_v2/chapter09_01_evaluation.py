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
    add_paragraph(document, '📊 技术位置解读：', bold=True)
    add_paragraph(document, f'• 当前价格：{current_price:.2f}元')
    add_paragraph(document, f'• 相对MA20：{((current_price/ma20 - 1) * 100):+.2f}%')
    add_paragraph(document, f'• 相对MA120：{((current_price/ma120 - 1) * 100):+.2f}%')
    add_paragraph(document, f'• 技术位置：{price_eval}')
    add_paragraph(document, '')

    # ==================== 9.1.2 相对估值分析汇总 ====================
    add_title(document, '9.1.2 相对估值分析汇总', level=3)

    add_paragraph(document, '本节汇总相对估值分析的核心结果，包括PE、PB等估值指标与行业的对比。')
    add_paragraph(document, '')

    # 相对估值数据（从第二章结果获取）
    current_pe = market_data.get('pe_ratio', 0)
    current_pb = market_data.get('pb_ratio', 0)
    industry_pe = market_data.get('industry_pe', 0)
    industry_pb = market_data.get('industry_pb', 0)

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
        add_paragraph(document, '📊 相对估值解读：', bold=True)
        add_paragraph(document, f'• PE相对行业：{pe_premium:+.1f}%（{pe_eval}）')
        if current_pb > 0 and industry_pb > 0:
            add_paragraph(document, f'• PB相对行业：{pb_premium:+.1f}%')
        add_paragraph(document, '')

    # ==================== 9.1.3 DCF估值分析汇总 ====================
    add_title(document, '9.1.3 DCF估值分析汇总', level=3)

    add_paragraph(document, '本节汇总DCF绝对估值分析的结果，评估公司的内在价值。')
    add_paragraph(document, '')

    # DCF估值数据（从第三章结果获取）
    intrinsic_value = context['results'].get('intrinsic_value', 0)
    if intrinsic_value > 0:
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
        add_paragraph(document, '📊 DCF估值解读：', bold=True)
        add_paragraph(document, f'• 内在价值：{intrinsic_value:.2f}元')
        add_paragraph(document, f'• 相对当前价：{dcf_discount:+.1f}%（{dcf_eval}）')
        add_paragraph(document, '')

    # ==================== 9.1.4 蒙特卡洛模拟结果汇总 ====================
    add_title(document, '9.1.4 蒙特卡洛模拟结果汇总', level=3)

    add_paragraph(document, '本节汇总蒙特卡洛模拟的核心结果，包括120日窗口的盈利概率和预期收益率。')
    add_paragraph(document, '')

    # 蒙特卡洛模拟结果（从第五章结果获取）
    mc_results = context['results'].get('mc_results', {})
    profit_prob_120d = mc_results.get('profit_prob_120d', 0.5)
    mean_return_120d = mc_results.get('mean_return_120d', 0.0)
    median_return_120d = mc_results.get('median_return_120d', 0.0)

    mc_summary_data = [
        ['模拟参数', '值'],
        ['模拟窗口', '120日（半年）'],
        ['模拟次数', '10,000次'],
        ['年化波动率', f'{mc_volatility*100:.2f}%'],
        ['年化漂移率', f'{mc_drift*100:+.2f}%'],
        ['', ''],
        ['模拟结果', ''],
        ['盈利概率', f'{profit_prob_120d*100:.1f}%'],
        ['平均收益率', f'{mean_return_120d*100:+.2f}%'],
        ['中位数收益率', f'{median_return_120d*100:+.2f}%'],
        ['风险评级', risk_rating]
    ]

    add_table_data(document, mc_summary_data[0], mc_summary_data[1:])

    add_paragraph(document, '')

    # ==================== 9.1.5 情景分析汇总 ====================
    add_title(document, '9.1.5 情景分析汇总', level=3)

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

    # ==================== 9.1.6 VaR风险度量汇总 ====================
    add_title(document, '9.1.6 VaR风险度量汇总', level=3)

    add_paragraph(document, '本节汇总VaR（Value at Risk）风险度量的结果，评估在不同置信水平下的最大损失。')
    add_paragraph(document, '')

    var_summary_data = [
        ['风险指标', '值', '说明'],
        ['95% VaR', f'{var_95*100:.1f}%', '95%置信度下的最大损失'],
        ['99% VaR', f'{var_99*100:.1f}%', '99%置信度下的最大损失'],
        ['风险评级', risk_rating, '综合风险评估']
    ]

    add_table_data(document, var_summary_data[0], var_summary_data[1:])

    add_paragraph(document, '')
    add_paragraph(document, '📊 VaR解读：', bold=True)
    add_paragraph(document, f'• 95% VaR：{var_95*100:.1f}%，意味着有95%的概率损失不超过此值')
    add_paragraph(document, f'• 99% VaR：{var_99*100:.1f}%，意味着有99%的概率损失不超过此值')
    add_paragraph(document, f'• 风险评级：{risk_rating}')
    add_paragraph(document, '')

    # 返回更新后的context
    return context
