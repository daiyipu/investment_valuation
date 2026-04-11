"""
第九章：风控建议与风险提示

本章节从风险控制角度，给出盈亏平衡分析、核心指标汇总、最终报价建议和全面的风险提示。
基于保守原则，确保投资决策在合理风险可控范围内进行。
"""

import sys
import os
from datetime import datetime
from docx.shared import Inches

# 添加路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)

from module_utils import add_title, add_paragraph, add_table_data, add_image, add_section_break
from module_utils import generate_break_even_chart, generate_market_turnover_chart
import chapter09_01_evaluation


def calculate_macro_environment_assessment(market_data, document, industry_cycle_override=None):
    """
    计算市场环境评估得分

    评估维度：
    1. 货币政策（15%权重）：紧缩(40分)/稳健(60分)/适度宽松(80分)/宽松(100分)
    2. 财政政策（15%权重）：紧缩(50分)/稳健(70分)/积极(90分)
    3. 行业发展周期（30%权重）：成长(100分)/成熟(80分)/幼稚(60分)/衰退(40分) - 支持参数指定
    4. 二级市场活跃度（40%权重）：基于历史分位数

    Args:
        market_data: 市场数据字典
        document: Word文档对象（用于添加表格）
        industry_cycle_override: 行业发展周期覆盖参数，可选值："成长"(100)/"成熟"(80)/"幼稚"(60)/"衰退"(40)
                                  如果不指定则默认为"幼稚"(60分)

    Returns:
        dict: 包含total_score, assessment_level, assessment_description等
    """
    import numpy as np

    # ==================== 1. 货币政策评估（15%权重）====================
    # 这里简化处理：基于市场利率和M2增速评估
    # 实际应用中可以根据最新宏观经济数据调整
    current_rdr = market_data.get('risk_free_rate', 0.03)  # 无风险利率作为代理指标

    # 简化的货币政策评分逻辑
    if current_rdr < 0.025:
        monetary_policy_score = 100  # 宽松
        monetary_policy_status = "宽松"
    elif current_rdr < 0.035:
        monetary_policy_score = 80  # 适度宽松
        monetary_policy_status = "适度宽松"
    elif current_rdr < 0.045:
        monetary_policy_score = 60  # 稳健
        monetary_policy_status = "稳健"
    else:
        monetary_policy_score = 40  # 紧缩
        monetary_policy_status = "紧缩"

    # ==================== 2. 财政政策评估（15%权重）====================
    # 基于财政赤字率和基建投资等指标
    # 这里简化处理，默认为"积极"
    fiscal_policy_score = 90  # 积极
    fiscal_policy_status = "积极"

    # ==================== 3. 行业发展周期评估（30%权重）====================
    # 支持参数指定行业发展周期，否则默认为幼稚期（60分）
    if industry_cycle_override is not None:
        # 使用参数指定的行业发展周期
        cycle_mapping = {
            "成长": 100,
            "成熟": 80,
            "幼稚": 60,
            "衰退": 40
        }
        industry_cycle_score = cycle_mapping.get(industry_cycle_override, 60)
        industry_cycle_status = industry_cycle_override
        # 添加说明：使用参数指定
        industry_cycle_note = f"（参数指定）"
    else:
        # 基于行业PE和行业估值水平评估
        industry_pe = market_data.get('industry_pe', 30)
        current_pe = market_data.get('pe_ratio', 30)

        # 行业相对估值：当前PE / 行业PE
        if industry_pe > 0:
            industry_relative_valuation = current_pe / industry_pe
        else:
            industry_relative_valuation = 1.0

        # 基于相对估值判断行业发展周期
        if industry_relative_valuation < 0.7:
            industry_cycle_score = 100  # 成长期（低估）
            industry_cycle_status = "成长"
        elif industry_relative_valuation < 0.9:
            industry_cycle_score = 80  # 成熟期（合理）
            industry_cycle_status = "成熟"
        elif industry_relative_valuation < 1.2:
            industry_cycle_score = 60  # 幼稚期（略高估）
            industry_cycle_status = "幼稚"
        else:
            industry_cycle_score = 40  # 衰退期（高估）
            industry_cycle_status = "衰退"

        # 如果无法判断，给予幼稚期默认分数60分
        if industry_pe == 0 or current_pe == 0:
            industry_cycle_score = 60
            industry_cycle_status = "幼稚"
            industry_cycle_note = f"（默认）"
        else:
            industry_cycle_note = f"（自动判断）"

    # ==================== 4. 二级市场活跃度评估（40%权重）====================
    # 基于历史换手率分位数（最近5年）

    # 获取市场换手率数据
    current_turnover_percentile = None
    if 'market_turnover' in market_data and market_data['market_turnover'] is not None:
        turnover_info = market_data['market_turnover']
        current_turnover_percentile = turnover_info.get('historical_percentile', 50)

    # 根据历史分位数确定活跃度等级和得分
    if current_turnover_percentile is not None:
        if current_turnover_percentile < 20:
            market_activity_score = 20
            market_activity_status = "不活跃"
        elif current_turnover_percentile < 40:
            market_activity_score = 40
            market_activity_status = "较不活跃"
        elif current_turnover_percentile < 60:
            market_activity_score = 60
            market_activity_status = "平稳"
        elif current_turnover_percentile < 80:
            market_activity_score = 80
            market_activity_status = "较活跃"
        else:
            market_activity_score = 100
            market_activity_status = "活跃"

        market_activity_note = f"（历史分位数{current_turnover_percentile:.1f}%）"
    else:
        # 如果没有换手率数据，使用默认值
        market_activity_score = 60
        market_activity_status = "平稳"
        market_activity_note = "（默认值，换手率数据不可用）"

    # ==================== 计算加权总分 ====================
    weights = {
        'monetary_policy': 0.15,
        'fiscal_policy': 0.15,
        'industry_cycle': 0.30,
        'market_activity': 0.40
    }

    weighted_scores = {
        'monetary_policy': monetary_policy_score * weights['monetary_policy'],
        'fiscal_policy': fiscal_policy_score * weights['fiscal_policy'],
        'industry_cycle': industry_cycle_score * weights['industry_cycle'],
        'market_activity': market_activity_score * weights['market_activity']
    }

    total_score = sum(weighted_scores.values())

    # ==================== 评估等级 ====================
    if total_score >= 90:
        assessment_level = "积极"
        assessment_description = "宏观环境良好，市场情绪乐观"
    elif total_score >= 80:
        assessment_level = "适度"
        assessment_description = "宏观环境较好，市场情绪稳定"
    elif total_score >= 70:
        assessment_level = "稳健"
        assessment_description = "宏观环境中性，市场情绪谨慎"
    elif total_score >= 60:
        assessment_level = "偏悲观"
        assessment_description = "宏观环境偏弱，市场情绪低迷"
    else:
        assessment_level = "悲观"
        assessment_description = "宏观环境较差，市场情绪悲观"

    # ==================== 生成评估表格 ====================
    assessment_data = [
        ['评估维度', '当前状态', '得分', '权重', '加权得分', '说明'],
        [
            '货币政策',
            monetary_policy_status,
            f'{monetary_policy_score:.0f}',
            '15%',
            f'{weighted_scores["monetary_policy"]:.1f}',
            '紧缩(40分)/稳健(60分)/适度宽松(80分)/宽松(100分)'
        ],
        [
            '财政政策',
            fiscal_policy_status,
            f'{fiscal_policy_score:.0f}',
            '15%',
            f'{weighted_scores["fiscal_policy"]:.1f}',
            '紧缩(50分)/稳健(70分)/积极(90分)'
        ],
        [
            '行业发展周期',
            f'{industry_cycle_status}{industry_cycle_note if "industry_cycle_note" in locals() else ""}',
            f'{industry_cycle_score:.0f}',
            '30%',
            f'{weighted_scores["industry_cycle"]:.1f}',
            '成长(100分)/成熟(80分)/幼稚(60分)/衰退(40分) - 默认幼稚期，可参数指定'
        ],
        [
            '二级市场活跃度',
            market_activity_status,
            f'{market_activity_score:.1f}',
            '40%',
            f'{weighted_scores["market_activity"]:.1f}',
            f'基于市场换手率历史分位数{market_activity_note}'
        ],
        [
            '合计',
            '',
            '',
            '100%',
            f'{total_score:.1f}',
            assessment_level
        ]
    ]

    add_table_data(document, assessment_data[0], assessment_data[1:])

    # 返回评估结果
    return {
        'total_score': total_score,
        'assessment_level': assessment_level,
        'assessment_description': assessment_description,
        'weighted_scores': weighted_scores,
        'individual_scores': {
            'monetary_policy': monetary_policy_score,
            'fiscal_policy': fiscal_policy_score,
            'industry_cycle': industry_cycle_score,
            'market_activity': market_activity_score
        },
        'individual_statuses': {
            'monetary_policy': monetary_policy_status,
            'fiscal_policy': fiscal_policy_status,
            'industry_cycle': industry_cycle_status,
            'market_activity': market_activity_status
        }
    }


def generate_chapter(context):
    """
    生成第九章和第十章：综合评估 + 投资建议与风险提示

    参数说明：
    - context: 包含所有必要数据的上下文字典
    """
    # 从context中提取变量
    document = context['document']
    project_params = context['project_params']
    market_data = context['market_data']
    IMAGES_DIR = context['IMAGES_DIR']

    # 从context['results']中获取其他章节的计算结果
    all_scenarios_for_appendix = context['results'].get('all_scenarios', [])
    intrinsic_value = context.get('intrinsic_value', 25.0)
    discount_premium = context['results'].get('discount_premium', 0.0)
    issue_type = project_params.get('issue_type', '竞价')

    # 从第五章获取蒙特卡洛结果
    mc_results = context['results'].get('mc_results', {})
    profit_prob = mc_results.get('profit_prob_120d', 0.5)
    mean_return = mc_results.get('mean_return_120d', 0.0)

    # 从第八章获取VaR结果
    var_95 = context['results'].get('var_95', 0.5)
    var_99 = context['results'].get('var_99', 0.7)

    # 其他参数（使用默认值或从context获取）
    total_score = context['results'].get('total_score', 50)
    ma20_mc = market_data.get('ma_20', project_params['current_price'])

    # 蒙特卡洛模拟参数（用于风险评估）
    mc_volatility_120d = market_data.get('volatility_120d', market_data.get('volatility', 0.30))
    mc_drift_120d = market_data.get('annual_return_120d', market_data.get('drift', 0.08))

    # 配置常量（使用context中的IMAGES_DIR）
    # IMAGES_DIR已在context中定义

    # ==================== 第九章标题 ====================
    add_title(document, '九、风控建议与风险提示', level=1)

    add_paragraph(document, '本章节从风险控制角度，给出盈亏平衡分析、核心指标汇总、最终报价建议和全面的风险提示。')
    add_paragraph(document, '基于保守原则，确保投资决策在合理风险可控范围内进行。')

    # ==================== 9.1 综合评估汇总 ====================
    # 生成9.1节的内容（通过调用chapter09_01_evaluation模块）
    print("\n 生成第九章第一节：综合评估汇总...")
    context = chapter09_01_evaluation.generate_chapter(context)

    # ==================== 9.2 盈亏平衡分析 ====================
    add_title(document, '9.2 盈亏平衡分析', level=2)
    add_paragraph(document, '本节通过盈亏平衡分析，评估在不同目标收益率和不同退出周期下的安全边际。')

    # ==================== 9.2.1 基于当前锁定期的盈亏平衡分析 ====================
    add_title(document, '9.2.1 基于当前锁定期的盈亏平衡分析', level=3)

    add_paragraph(document, '本节基于当前锁定期，分析不同目标收益率下的盈亏平衡价格。')

    # 计算不同收益率下的盈亏平衡价格
    import numpy as np

    # 定义要显示的收益率档位：10%, 15%, 20%, 25%, 30%, 35%, 40%, 45%
    display_returns = np.array([0.10, 0.15, 0.20, 0.25, 0.30, 0.35, 0.40, 0.45])

    issue_price = project_params['issue_price']
    issue_price_eval = issue_price  # 添加此变量供后续使用
    lockup_years = project_params['lockup_period'] / 12
    current_price_eval = project_params['current_price']

    # 检查发行日是否确定，用于区分"发行日价格"和"发行日价格（拟）"
    is_fixed_issue_date = project_params.get('invitation_date_fixed', False)
    if is_fixed_issue_date:
        issue_date = project_params.get('issue_date', '')
        price_label = '发行日价格'
        price_description = f'发行日价格（{issue_date}）'
    else:
        price_label = '发行日价格（拟）'
        price_description = '发行日价格（拟）'

    # 为每个收益率档位计算盈亏平衡价格
    break_even_prices = []
    for target_return in display_returns:
        # 使用复利计算：发行价 × (1 + 年化收益率)^锁定期年数
        be_price = issue_price * (1 + target_return) ** lockup_years
        break_even_prices.append(be_price)

    # 生成表格数据
    be_data = []
    for ret, price in zip(display_returns, break_even_prices):
        # 计算实际需要的收益率（锁定期内）
        actual_return = (price - issue_price) / issue_price
        # 计算离目标的差距：(盈亏平衡价 / 发行日价格) - 1，以发行日价格为基准
        gap_to_target = (price / current_price_eval - 1) * 100
        status = "" if gap_to_target > 0 else ""
        be_data.append([f'{ret*100:.0f}%', f'{actual_return*100:.1f}%', f'{price:.2f}', f'{gap_to_target:+.1f}%', status])

    add_table_data(document, ['期望年化收益率', '实际需要收益率', '盈亏平衡价格(元)', '离目标差距', '状态'], be_data)

    add_paragraph(document, '说明：年化收益率是基于全年计算的收益率，实际需要收益率是锁定期内需要的价格涨幅。')
    add_paragraph(document, f'离目标差距：基于{price_description}计算，正值表示需要上涨才能达到目标，负值表示{price_label}已超过目标有安全边际。')
    add_paragraph(document, '')

    add_paragraph(document, '盈亏平衡分析结论：')
    add_paragraph(document, f'• 发行价格: {issue_price:.2f} 元/股')
    add_paragraph(document, f'• {price_label}: {current_price_eval:.2f} 元/股')
    add_paragraph(document, f'• 锁定期: {project_params["lockup_period"]}个月（{lockup_years:.2f}年）')

    # 20%收益率对应的索引是2（display_returns中：10%, 15%, 20%, ...）
    be_20 = break_even_prices[2]
    add_paragraph(document, f'• 20%年化收益率要求下盈亏平衡价: {be_20:.2f} 元')

    # 以发行日价格为基准计算差距
    gap_to_target = (be_20 / current_price_eval - 1) * 100
    if gap_to_target > 0:
        add_paragraph(document, f'•  {price_label}需要上涨{gap_to_target:.1f}%才能达到20%收益率目标')
    else:
        add_paragraph(document, f'•  {price_label}已超过20%收益率目标{abs(gap_to_target):.1f}%，有安全边际')

    
    # ==================== 9.2.2 考虑退出周期的盈亏平衡分析 ====================
    add_title(document, '9.2.2 考虑退出周期的盈亏平衡分析', level=3)

    add_paragraph(document, '本节考虑不同的退出周期（锁定期），计算在不同周期下的盈亏平衡价格。')
    add_paragraph(document, '分析考虑资金的时间成本，年化资金成本统一按4%计算。')
    add_paragraph(document, '注：盈亏平衡价使用复利计算：发行价 × (1 + 年化收益率)^锁定期年数')

    # 期望收益率和资金成本
    net_return = 0.08  # 净收益率（期望收益率）年化8%，已扣除资金成本
    cost_of_capital = 0.04  # 资金成本年化4%
    total_return = net_return + cost_of_capital  # 总收益率要求12%

    add_paragraph(document, ' 计算参数：', bold=True)
    add_paragraph(document, f'• 期望收益率：{net_return*100:.0f}%（年化，已扣除资金成本）')
    add_paragraph(document, f'• 资金成本：{cost_of_capital*100:.0f}%（年化）')
    add_paragraph(document, f'• 总收益率要求：{total_return*100:.0f}%（年化）')
    add_paragraph(document, f'• 当前发行价：{issue_price:.2f} 元/股')
    add_paragraph(document, f'• 当前价格：{current_price_eval:.2f} 元/股')
    add_paragraph(document, '')

    # 按照3个月一期，从6个月到24个月
    lockup_periods = [6, 9, 12, 15, 18, 21, 24]  # 单位：月
    period_analysis = []

    for months in lockup_periods:
        years = months / 12

        # 计算考虑资金成本后的盈亏平衡价
        # 公式：盈亏平衡价 = 发行价 × (1 + 总收益率要求)^年数
        # 总收益率要求 = 净收益率（期望收益率）+ 资金成本

        # 方案1：考虑资金成本和时间价值
        be_price_with_cost = issue_price * (1 + total_return) ** years

        # 方案2：传统方法（仅考虑净收益率的时间价值）
        # 传统方法下，盈亏平衡价应该考虑锁定期内的净收益增长
        be_price_traditional = issue_price * (1 + net_return) ** years

        # 距离当前价的差距
        distance_with_cost = (current_price_eval - be_price_with_cost) / current_price_eval * 100
        distance_traditional = (current_price_eval - be_price_traditional) / current_price_eval * 100

        period_analysis.append({
            'months': months,
            'years': years,
            'be_price_with_cost': be_price_with_cost,
            'be_price_traditional': be_price_traditional,
            'distance_with_cost': distance_with_cost,
            'distance_traditional': distance_traditional
        })

    # 生成表格数据
    period_data = [
        ['退出周期', '盈亏平衡价(考虑资金成本)', '盈亏平衡价(传统)', '距离当前价(考虑成本)', '距离当前价(传统)', '评估'],
    ]

    for item in period_analysis:
        months = item['months']
        distance_with_cost = item['distance_with_cost']
        distance_traditional = item['distance_traditional']

        # 评估状态
        if distance_with_cost > 0:
            eval_text = "安全边际充足"
        elif distance_with_cost > -5:
            eval_text = "安全边际有限"
        else:
            eval_text = "安全边际不足"

        period_data.append([
            f'{months}个月',
            f'{item["be_price_with_cost"]:.2f}元',
            f'{item["be_price_traditional"]:.2f}元',
            f'{distance_with_cost:+.1f}%',
            f'{distance_traditional:+.1f}%',
            eval_text
        ])

    add_table_data(document, period_data[0], period_data[1:])

    add_paragraph(document, ' 退出周期分析结论：', bold=True)
    add_paragraph(document, '说明：')
    add_paragraph(document, f'• 考虑资金成本方法：盈亏平衡价 = 发行价 × (1 + {total_return*100:.0f}% × 年数)')
    add_paragraph(document, f'• 传统方法：盈亏平衡价 = 发行价 × (1 + {net_return*100:.0f}% × 年数)（仅考虑期望收益率的时间价值）')
    add_paragraph(document, f'• 总收益率要求{total_return*100:.0f}% = 期望收益率{net_return*100:.0f}% + 资金成本{cost_of_capital*100:.0f}%')

    # 分析不同退出周期的影响
    add_paragraph(document, ' 退出周期影响分析：', bold=True)

    # 找出最佳退出周期
    best_period = max(period_analysis, key=lambda x: x['distance_with_cost'])
    add_paragraph(document, f'• 最佳退出周期：{best_period["months"]}个月（{best_period["years"]:.2f}年）')
    add_paragraph(document, f'  盈亏平衡价：{best_period["be_price_with_cost"]:.2f}元（考虑资金成本）')
    add_paragraph(document, f'  距离当前价：{best_period["distance_with_cost"]:+.1f}%')

    if best_period['distance_with_cost'] > 0:
        add_paragraph(document, f'   在{best_period["months"]}个月退出时，当前价格{distance_with_cost:+.1f}%高于盈亏平衡价')
    else:
        add_paragraph(document, f'   在{best_period["months"]}个月退出时，当前价格{distance_with_cost:.1f}%低于盈亏平衡价')

    add_paragraph(document, '')

    # 对比当前锁定期
    current_lockup_months = project_params['lockup_period']
    current_result = next((item for item in period_analysis if item['months'] == current_lockup_months), None)

    if current_result:
        add_paragraph(document, f'• 当前锁定期（{current_lockup_months}个月）：')
        add_paragraph(document, f'  盈亏平衡价（考虑资金成本）：{current_result["be_price_with_cost"]:.2f}元')
        add_paragraph(document, f'  距离当前价：{current_result["distance_with_cost"]:+.1f}%')

        if current_result['distance_with_cost'] > 0:
            add_paragraph(document, f'   当前价格高于盈亏平衡价{current_result["distance_with_cost"]:.1f}%，具有安全边际')
        else:
            add_paragraph(document, f'   当前价格低于盈亏平衡价{abs(current_result["distance_with_cost"]):.1f}%，安全边际不足')
    add_paragraph(document, ' 投资建议：', bold=True)
    add_paragraph(document, '• 退出周期越长，资金成本越高，要求的盈亏平衡价也越高')
    add_paragraph(document, '• 建议选择退出周期较短（6-12个月）的项目，以降低资金成本压力')
    add_paragraph(document, '• 如果当前价格低于考虑资金成本的盈亏平衡价，建议要求更高的折价率')

    # 生成并插入盈亏平衡价格敏感性图表
    break_even_chart_path = os.path.join(IMAGES_DIR, '09_break_even_analysis.png')
    generate_break_even_chart(issue_price, current_price_eval, project_params['lockup_period'], break_even_chart_path)
    add_paragraph(document, '图表 9.1: 盈亏平衡价格敏感性曲线')
    add_image(document, break_even_chart_path, width=Inches(6))

    # ==================== 9.3 报价方案建议 ====================
    add_title(document, '9.3 报价方案建议', level=2)
    add_paragraph(document, '本节提供不同目标收益率下的报价建议，帮助投资者根据风险偏好选择合适的报价方案。')

    # ==================== 市场环境评估 ====================
    add_paragraph(document, '【市场环境评估】', bold=True)
    add_paragraph(document, '在进行报价方案建议之前，首先评估当前宏观环境，为情景选择提供参考依据。')
    add_paragraph(document, '')

    # 计算宏观环境评分
    macro_assessment = calculate_macro_environment_assessment(market_data, document)

    # 保存市场环境评估结果到context，供后续情景选择使用
    context['macro_environment_assessment'] = macro_assessment

    # 显示市场环境评估结论
    total_score = macro_assessment['total_score']
    assessment_level = macro_assessment['assessment_level']

    add_paragraph(document, f'综合评分：{total_score:.1f}分 | 评估等级：{assessment_level}')

    # 根据评估等级给出情景选择建议
    if total_score >= 90:
        add_paragraph(document, f' 宏观环境"积极"（{total_score:.1f}分）：建议溢价率和波动率选择第一档（乐观），漂移率结合公司基本面选择')
    elif total_score >= 80:
        add_paragraph(document, f' 宏观环境"适度"（{total_score:.1f}分）：建议溢价率和波动率选择第二档（中性偏乐观），漂移率需结合公司基本面判断')
    elif total_score >= 70:
        add_paragraph(document, f' 宏观环境"稳健"（{total_score:.1f}分）：建议溢价率和波动率选择第三档（中性），漂移率结合公司基本面选择偏保守一档')
    elif total_score >= 60:
        add_paragraph(document, f'⚠ 宏观环境"偏悲观"（{total_score:.1f}分）：建议溢价率和波动率选择第四档（谨慎），漂移率结合公司基本面选择保守档')
    else:
        add_paragraph(document, f' 宏观环境"悲观"（{total_score:.1f}分）：建议溢价率和波动率选择第五档（极度保守），漂移率结合公司基本面选择最保守档或谨慎参与')

    # 添加公司基本面判断说明
    add_paragraph(document, '')
    add_paragraph(document, '【公司基本面判断说明】', bold=True)
    add_paragraph(document, '漂移率选择需结合公司基本面判断（经营周期）：')
    add_paragraph(document, '• 优质期（业绩增长稳定、行业地位领先）：漂移率范围 +10% 至 +30%')
    add_paragraph(document, '• 成熟期（业绩稳定、行业地位稳固）：漂移率范围 0% 至 +10%')
    add_paragraph(document, '• 衰退期（业绩下滑、行业地位下降）：漂移率范围 -10% 至 0%')
    add_paragraph(document, '• 极度保守期：漂移率范围 -30% 至 -10%')
    add_paragraph(document, '后续将根据经营周期对应的漂移率范围，筛选符合条件的情景。')
    # 参数档位说明
    add_paragraph(document, '【参数档位说明】', bold=True)
    add_paragraph(document, '溢价率档位（从优到差）：')
    add_paragraph(document, '• 第一档：0%（无折价，最高报价）')
    add_paragraph(document, '• 第二档：-5%（小幅折价）')
    add_paragraph(document, '• 第三档：-10%（中等折价）')
    add_paragraph(document, '• 第四档：-15%（较大折价）')
    add_paragraph(document, '• 第五档：-20%（最大折价，最低报价）')

    add_paragraph(document, '波动率档位（从高到低）：')
    add_paragraph(document, '• 第一档：40%-50%（高波动，风险较高）')
    add_paragraph(document, '• 第二档：30%-40%（中高波动）')
    add_paragraph(document, '• 第三档：20%-30%（中低波动）')
    add_paragraph(document, '• 第四档：10%-20%（低波动，风险较低）')
    add_paragraph(document, '')
    add_paragraph(document, '注：上述溢价率建议基于市场环境评估，具体报价还需结合项目基本面和市场情况综合判断。')
    # 9.3.1 市场环境评估
    add_title(document, '9.3.1 市场环境评估', level=3)
    add_paragraph(document, '在制定报价方案前，先评估当前的宏观环境，包括货币政策与财政政策、行业发展周期、二级市场活跃度三个维度。')

    # 计算二级市场活跃度（使用换手率）
    print(f"正在读取市场换手率数据...")

    # 尝试从market_data中读取
    if 'market_turnover' in market_data and market_data['market_turnover'] is not None:
        turnover_info = market_data['market_turnover']
        weighted_turnover = turnover_info.get('current_turnover', 1.99)
        current_turnover_percentile = turnover_info.get('historical_percentile', 50)
        historical_count = turnover_info.get('historical_count', 0)

        print(f" 成功读取市场换手率数据：")
        print(f"   当前换手率={weighted_turnover:.2f}%")
        print(f"   历史分位数={current_turnover_percentile:.1f}%（基于{historical_count}个交易日）")
    else:
        print(f" 市场换手率数据不可用，使用默认值")
        weighted_turnover = 1.99
        current_turnover_percentile = 50
        historical_count = 0
        turnover_info = {}

    # 根据历史分位数确定活跃度等级和得分
    if current_turnover_percentile < 20:
        market_activity_level = "不活跃"
        market_activity_score = 20
        market_activity_desc = "历史分位数<20%（5年）"
    elif current_turnover_percentile < 40:
        market_activity_level = "较不活跃"
        market_activity_score = 40
        market_activity_desc = f"历史分位数20%-40%（5年，当前{current_turnover_percentile:.1f}%）"
    elif current_turnover_percentile < 60:
        market_activity_level = "平稳"
        market_activity_score = 60
        market_activity_desc = f"历史分位数40%-60%（5年，当前{current_turnover_percentile:.1f}%）"
    elif current_turnover_percentile < 80:
        market_activity_level = "较活跃"
        market_activity_score = 80
        market_activity_desc = f"历史分位数60%-80%（5年，当前{current_turnover_percentile:.1f}%）"
    else:
        market_activity_level = "活跃"
        market_activity_score = 100
        market_activity_desc = f"历史分位数>80%（5年，当前{current_turnover_percentile:.1f}%）"

    current_percentile = current_turnover_percentile

    # 市场环境评估表格
    add_paragraph(document, ' 宏观环境三维度评估：', bold=True)

    # 货币政策和财政政策（固定值）
    monetary_policy = '适度宽松'
    monetary_policy_score = 80
    monetary_policy_weight = 0.15  # 15%
    monetary_policy_weighted = monetary_policy_score * monetary_policy_weight

    fiscal_policy = '积极'
    fiscal_policy_score = 90
    fiscal_policy_weight = 0.15  # 15%
    fiscal_policy_weighted = fiscal_policy_score * fiscal_policy_weight

    # 行业发展周期（初始值：幼稚期60分，可根据实际情况调整）
    industry_cycle = '幼稚期'
    industry_cycle_score = 60
    industry_cycle_weight = 0.30  # 30%
    industry_cycle_weighted = industry_cycle_score * industry_cycle_weight

    macro_env_data = [
        ['评估维度', '当前状态', '得分', '权重', '加权得分', '说明'],
        # 1. 货币政策（权重15%）
        ['货币政策', monetary_policy, f'{monetary_policy_score}', '15%', f'{monetary_policy_weighted:.1f}',
         '紧缩(40分)/稳健(60分)/适度宽松(80分)/宽松(100分)'],
        # 2. 财政政策（权重15%）
        ['财政政策', fiscal_policy, f'{fiscal_policy_score}', '15%', f'{fiscal_policy_weighted:.1f}',
         '紧缩(50分)/稳健(70分)/积极(90分)'],
        # 3. 行业发展周期（权重30%，初始值幼稚期60分）
        ['行业发展周期', industry_cycle, f'{industry_cycle_score}', '30%', f'{industry_cycle_weighted:.1f}',
         '成长(100分)/成熟(80分)/幼稚(60分)/衰退(40分) - 可根据实际情况调整'],
        # 4. 二级市场活跃度（权重40%，自动计算）
        ['二级市场活跃度', market_activity_level, f'{market_activity_score}', '40%', f'{market_activity_score * 0.4:.1f}',
         f'{market_activity_desc}，基于最近5年历史换手率计算'],
    ]

    # 分离表头和数据
    macro_env_headers = macro_env_data[0]
    macro_env_table_data = macro_env_data[1:]
    add_table_data(document, macro_env_headers, macro_env_table_data)

    # 计算宏观环境总分
    macro_total_score = monetary_policy_weighted + fiscal_policy_weighted + industry_cycle_weighted + (market_activity_score * 0.4)

    # 将总分保存到context中，供后续筛选使用
    context['macro_score'] = macro_total_score

    add_paragraph(document, f'• 宏观环境总分：{macro_total_score:.1f}分（满分100分）')

    # 评估说明
    add_paragraph(document, '评估说明：', bold=True)
    add_paragraph(document, f'• 宏观环境总分：{macro_total_score:.1f}分（满分100分）')
    add_paragraph(document, '• 评估周期：基于最近5年历史数据计算分位数')
    add_paragraph(document, '• 权重分配：货币政策(15%) + 财政政策(15%) + 行业发展周期(30%) + 二级市场活跃度(40%)')
    add_paragraph(document, '• 得分标准：')
    add_paragraph(document, '  - 货币政策：紧缩(40分)、稳健(60分)、适度宽松(80分)、宽松(100分)')
    add_paragraph(document, '  - 财政政策：紧缩(50分)、稳健(70分)、积极(90分)')
    add_paragraph(document, '  - 行业发展周期：成长(100分)、成熟(80分)、幼稚(60分)、衰退(40分)')
    add_paragraph(document, '  - 二级市场活跃度：根据历史分位数自动计算（20-100分）')

    # 二级市场活跃度详细说明
    add_paragraph(document, ' 二级市场活跃度详情：', bold=True)
    add_paragraph(document, f'• 当前状态：{market_activity_level}（得分{market_activity_score}分）')
    add_paragraph(document, f'• 当前换手率：{weighted_turnover:.2f}%（120日中位数）')
    add_paragraph(document, f'• 历史分位数：{current_percentile:.1f}%（最近{historical_count}个交易日）')
    add_paragraph(document, '• 数据来源：深圳市场换手率（代表全市场，tushare daily_info接口）')
    add_paragraph(document, '')

    # 生成并添加换手率曲线图
    if 'market_turnover' in market_data and market_data['market_turnover'] is not None:
        turnover_chart_path = os.path.join(IMAGES_DIR, '09_market_turnover_history.png')
        print(f"正在生成市场换手率曲线图...")
        chart_path = generate_market_turnover_chart(market_data['market_turnover'], turnover_chart_path)
        if chart_path and os.path.exists(chart_path):
            add_paragraph(document, '图表 9.3: 市场换手率历史走势')
            add_image(document, chart_path, width=Inches(6.0))
            add_paragraph(document, '')
            add_paragraph(document, '图表说明：')
            add_paragraph(document, f'• 蓝色曲线：历史换手率走势（共{historical_count}个交易日）')
            add_paragraph(document, f'• 红色虚线：当前120日中位数（{weighted_turnover:.2f}%），代表近期市场活跃度水平')
            add_paragraph(document, f'• 绿色点线：最新一日换手率（{turnover_info.get("latest_turnover", 0):.2f}%），反映当日市场情绪')
            add_paragraph(document, f'• 当前分位数：{current_percentile:.1f}%，表示当前活跃度在历史中的相对位置')

    # 活跃度对定增的影响
    add_paragraph(document, ' 活跃度对定增的影响：', bold=True)
    if market_activity_score >= 80:
        add_paragraph(document, f'•  市场较活跃，流动性充足，有利于定增项目发行和退出')
        add_paragraph(document, f'• 投资者情绪较高，可适当提高报价预期')
    elif market_activity_score >= 60:
        add_paragraph(document, f'•  市场平稳，流动性适中，定增项目正常发行')
        add_paragraph(document, f'• 建议按照标准折价率报价')
    else:
        add_paragraph(document, f'•  市场不活跃，流动性偏紧，定增发行难度增加')
        add_paragraph(document, f'• 建议要求更高折价率或谨慎参与')

    add_paragraph(document, '')

    # 市场环境评分参数映射表（用于参数构造场景585种）
    add_paragraph(document, '市场环境评分参数映射表（参数构造场景）：', bold=True)
    add_paragraph(document, '本表定义了不同宏观评分下对应的漂移率和波动率参数范围，用于585种参数构造场景的筛选。')
    add_paragraph(document, '')

    # 创建映射表数据
    mapping_table_data = [
        ['80-100分', '市场环境较好', '25% - 35%', '40% - 50%', '高分环境：高漂移率，高波动率'],
        ['75-80分', '市场环境良好', '20% - 30%', '35% - 45%', '较高分环境：较高漂移率和波动率'],
        ['70-75分', '市场环境中上', '15% - 25%', '30% - 40%', '中上分环境：中等偏上参数'],
        ['65-70分', '市场环境中等', '10% - 20%', '25% - 35%', '中等分环境：中等参数'],
        ['60-65分', '市场环境中下', '5% - 15%', '20% - 30%', '中下分环境：中等偏下参数'],
        ['50-60分', '市场环境偏差', '0% - 10%', '18% - 28%', '较低分环境：较低参数'],
        ['40-50分', '市场环境较差', '-5% - 5%', '15% - 25%', '低分环境：低参数'],
        ['<40分', '市场环境很差', '-10% - 0%', '12% - 22%', '极低分环境：极低参数']
    ]

    add_table_data(document, ['宏观评分范围', '市场环境', '漂移率范围', '波动率范围', '说明'], mapping_table_data)
    add_paragraph(document, '')

    # 9.3.2 报价方案建议
    add_title(document, '9.3.2 报价方案建议', level=3)
    add_paragraph(document, '本节基于历史收益率和情景分析，提供多角度的报价建议。')

    # 定义市场环境评分映射函数（供整个9.3.2节使用）
    def get_drift_volatility_mapping(macro_score):
        """
        根据宏观评分获取对应的漂移率和波动率范围
        返回：(drift_min, drift_max, volatility_min, volatility_max)
        """
        if macro_score >= 80:
            # 高分环境：高漂移率，高波动率
            return (0.25, 0.35, 0.40, 0.50)
        elif macro_score >= 75:
            # 较高分环境
            return (0.20, 0.30, 0.35, 0.45)
        elif macro_score >= 70:
            return (0.15, 0.25, 0.30, 0.40)
        elif macro_score >= 65:
            return (0.10, 0.20, 0.25, 0.35)
        elif macro_score >= 60:
            return (0.05, 0.15, 0.20, 0.30)
        elif macro_score >= 50:
            return (0.00, 0.10, 0.18, 0.28)
        elif macro_score >= 40:
            return (-0.05, 0.05, 0.15, 0.25)
        else:
            # 低分环境：低漂移率，低波动率
            return (-0.10, 0.00, 0.12, 0.22)

    add_title(document, '9.3.2.1 反向推算最高报价', level=4)

    # 处理报价日信息
    # 优先使用项目参数中的具体报价日，如果没有则使用当前日（遇节假日往前推）
    if 'issue_date' in project_params and project_params['issue_date']:
        # 如果项目中有具体的报价日
        issue_date_input = project_params['issue_date']
        if isinstance(issue_date_input, str):
            issue_date_display = issue_date_input  # 已经是字符串格式
        else:
            issue_date_display = issue_date_input.strftime('%Y年%m月%d日')
        date_source = "指定报价日"
    else:
        # 如果没有具体报价日，使用当前日
        today = datetime.now()
        issue_date_display = today.strftime('%Y年%m月%d日')
        date_source = "当前日（遇节假日往前推）"

    add_paragraph(document, f'• 报价日期：{issue_date_display}（{date_source}）')

    # 计算参数
    target_return = 0.08  # 目标收益率8%
    lockup_years = project_params['lockup_period'] / 12

    # 获取市场环境评分
    macro_score = context.get('macro_score', 83.5)

    # 使用已有的映射函数获取波动率范围
    drift_min, drift_max, vol_min, vol_max = get_drift_volatility_mapping(macro_score)
    # 使用波动率范围的中值
    environment_volatility = (vol_min + vol_max) / 2

    current_price_eval = project_params['current_price']
    ma20_price = market_data.get('ma_20', current_price_eval)

    # 方法1：简单反推（保守）- 不考虑波动率
    # 使用复利：最高发行价 = 当前价格 / (1 + 目标收益率)^锁定期年数
    max_issue_price_simple = current_price_eval / (1 + target_return) ** lockup_years
    premium_to_current_simple = (max_issue_price_simple - current_price_eval) / current_price_eval * 100
    premium_to_ma20_simple = (max_issue_price_simple - ma20_price) / ma20_price * 100

    # 方法2：考虑市场环境波动率的反推（使用安全边际）
    safety_margin = 0.5 * environment_volatility ** 2
    adjusted_return = target_return + safety_margin
    # 使用复利：最高发行价 = 当前价格 / (1 + 调整后收益率)^锁定期年数
    max_issue_price_adjusted = current_price_eval / (1 + adjusted_return) ** lockup_years
    premium_to_current_adjusted = (max_issue_price_adjusted - current_price_eval) / current_price_eval * 100
    premium_to_ma20_adjusted = (max_issue_price_adjusted - ma20_price) / ma20_price * 100

    # 执行逻辑说明
    add_paragraph(document, '执行逻辑说明：', bold=True)
    add_paragraph(document, '基于目标收益率8%，使用以下逻辑反推最高可接受溢价率：')
    add_paragraph(document, f'1. 设定目标收益率：8%（年化）')
    add_paragraph(document, f'2. 市场环境评分：{macro_score:.1f}分，对应波动率范围：{vol_min*100:.0f}%-{vol_max*100:.0f}%（中值：{environment_volatility*100:.1f}%）')
    add_paragraph(document, f'3. 锁定期：{project_params["lockup_period"]}个月（{lockup_years:.2f}年）')
    add_paragraph(document, f'4. 报价日期：{issue_date_display}')
    add_paragraph(document, f'5. 参考价格：当前价格{current_price_eval:.2f}元（报价日价格），MA20价格{ma20_price:.2f}元')

    add_paragraph(document, '计算方法：', bold=True)
    add_paragraph(document, '• 方法1（保守）：简单反推，不考虑波动率')
    add_paragraph(document, f'  当前价格 = {current_price_eval:.2f}元（报价日价格）')
    add_paragraph(document, '  公式：最高发行价 = 当前价格 ÷ (1 + 目标收益率)^锁定期年数')
    add_paragraph(document, f'  最高发行价 = {current_price_eval:.2f} ÷ (1 + 8%)^{lockup_years:.2f} = {max_issue_price_simple:.2f}元')
    add_paragraph(document, f'  对应名义溢价率（相对MA20）：{premium_to_ma20_simple:+.2f}%（MA20={ma20_price:.2f}元）')
    add_paragraph(document, f'  对应实际溢价率（相对当前价）：{premium_to_current_simple:+.2f}%')
    add_paragraph(document, '')

    add_paragraph(document, '• 方法2（更保守）：考虑市场环境波动率安全边际')
    add_paragraph(document, f'  当前价格 = {current_price_eval:.2f}元（报价日价格）')
    add_paragraph(document, f'  市场环境波动率 = {environment_volatility*100:.1f}%（基于评分{macro_score:.1f}分的映射值）')
    add_paragraph(document, f'  安全边际 = 0.5 × 波动率² = 0.5 × ({environment_volatility*100:.1f}%)² = {safety_margin*100:.2f}%')
    add_paragraph(document, f'  调整后收益率 = 目标收益率 + 安全边际 = 8% + {safety_margin*100:.2f}% = {adjusted_return*100:.2f}%')
    add_paragraph(document, f'  最高发行价 = {current_price_eval:.2f} ÷ (1 + {adjusted_return*100:.2f}%)^{lockup_years:.2f} = {max_issue_price_adjusted:.2f}元')
    add_paragraph(document, f'  对应名义溢价率（相对MA20）：{premium_to_ma20_adjusted:+.2f}%（MA20={ma20_price:.2f}元）')
    add_paragraph(document, f'  对应实际溢价率（相对当前价）：{premium_to_current_adjusted:+.2f}%')
    add_paragraph(document, '')

    # 推荐方案对比表
    add_paragraph(document, '推荐方案对比：', bold=True)

    pricing_options = [
        ['计算方法', '最高发行价', '实际溢价率（相对报价日价格）', '名义溢价率（相对MA20）', '特点'],
        [
            '方法1：简单反推',
            f'{max_issue_price_simple:.2f}元',
            f'{premium_to_current_simple:+.2f}%',
            f'{premium_to_ma20_simple:+.2f}%',
            '不考虑波动率，相对激进'
        ],
        [
            '方法2：考虑波动率安全边际',
            f'{max_issue_price_adjusted:.2f}元',
            f'{premium_to_current_adjusted:+.2f}%',
            f'{premium_to_ma20_adjusted:+.2f}%',
            '更保守，推荐使用'
        ],
    ]

    add_table_data(document, pricing_options[0], pricing_options[1:])

    # 推荐结论
    add_paragraph(document, '')
    add_paragraph(document, '推荐结论：', bold=True)

    # 选择更保守的方法2
    recommended_price = max_issue_price_adjusted
    premium_to_ma20 = premium_to_ma20_adjusted  # 名义溢价率（相对MA20）
    premium_to_current = premium_to_current_adjusted  # 实际溢价率（相对报价日价格）

    add_paragraph(document, '• 推荐使用方法2（考虑市场环境波动率安全边际）')
    add_paragraph(document, f'• 建议最高报价：不高于{recommended_price:.2f}元/股')
    add_paragraph(document, f'• 名义溢价率（相对MA20）：{premium_to_ma20:+.2f}%')
    # 根据是否有确定报价日显示不同的描述
    if date_source == "指定报价日":
        add_paragraph(document, f'• 实际溢价率（相对报价日价格{current_price_eval:.2f}元）：{premium_to_current:+.2f}%')
    else:
        add_paragraph(document, f'• 实际溢价率（相对当前价格{current_price_eval:.2f}元，默认为报价日价格）：{premium_to_current:+.2f}%')
    add_paragraph(document, f'• 该溢价率已考虑市场环境评分{macro_score:.1f}分对应波动率的安全边际')
    add_paragraph(document, '')

    # 投资建议
    add_paragraph(document, ' 投资建议：', bold=True)

    # 根据是否有确定报价日显示不同的建议描述
    price_desc = f"报价日价格{current_price_eval:.2f}元" if date_source == "指定报价日" else f"当前价格{current_price_eval:.2f}元（默认为报价日价格）"

    if premium_to_ma20 <= -5:
        add_paragraph(document, f'•  建议报价 - 名义溢价率{premium_to_ma20:+.2f}%，实际溢价率{premium_to_current:+.2f}%（相对{price_desc}），安全边际充足，符合保守投资原则')
    elif premium_to_ma20 <= 0:
        add_paragraph(document, f'•  谨慎报价 - 名义溢价率{premium_to_ma20:+.2f}%，实际溢价率{premium_to_current:+.2f}%（相对{price_desc}），安全边际有限，建议折价至少5%')
    else:
        add_paragraph(document, f'•  不建议报价 - 名义溢价率{premium_to_ma20:+.2f}%为溢价发行，实际溢价率{premium_to_current:+.2f}%（相对{price_desc}），缺乏安全边际')
        add_paragraph(document, '  建议：等待市场时机或选择其他投资机会')

    add_paragraph(document, ' 注意事项：')
    add_paragraph(document, '• 本计算基于目标收益率8%，实际收益可能因市场波动而偏离')
    add_paragraph(document, '• 建议结合其他分析方法（如DCF估值、相对估值）综合判断')
    add_paragraph(document, '• 如果实际发行价高于推荐报价，建议谨慎参与或放弃')

    # ==================== 9.3.2.2 基于历史数据场景的筛选 ====================
    add_title(document, '9.3.2.2 基于历史数据场景的筛选', level=4)

    add_paragraph(document, '本节基于第六章的195种历史数据场景（市场指数、行业指数、行业PE、个股PE、DCF估值），通过严格筛选条件，从风险收益平衡角度推荐最优报价方案。')

    # 市场环境条件筛选
    add_paragraph(document, ' 市场环境条件筛选：', bold=True)

    # 获取宏观评分
    macro_score = context.get('macro_score', 60)

    # 根据宏观评分确定档位选择策略
    if macro_score >= 80:
        tier_level = "高档位"
        tier_selection = "选择各历史数据场景的高档位（如75%分位数、高估水平）"
        market_env_desc = f"市场宏观评分{macro_score}分（≥80分），市场环境较好"
    elif macro_score >= 60:
        tier_level = "中档位"
        tier_selection = "选择各历史数据场景的中档位（如50%分位数、中性水平）"
        market_env_desc = f"市场宏观评分{macro_score}分（60-80分），市场环境一般"
    else:
        tier_level = "低档位"
        tier_selection = "选择各历史数据场景的低档位（如25%分位数、低估水平）"
        market_env_desc = f"市场宏观评分{macro_score}分（<60分），市场环境较差"

    add_paragraph(document, f'• {market_env_desc}')
    add_paragraph(document, f'• 档位选择：{tier_selection}')
    add_paragraph(document, f'• 筛选逻辑：根据市场环境评分，从每个历史数据场景的多个档位中选择对应{tier_level}进行分析')
    add_paragraph(document, '')

    # 基础筛选条件
    add_paragraph(document, ' 基础筛选条件（必须全部满足）：', bold=True)
    add_paragraph(document, '1. 溢价率 ≤ -5%（折价至少5%，提供安全边际）')
    add_paragraph(document, '2. 预测中位数收益率 > 8%（年化收益达标）')
    add_paragraph(document, '3. 95% VaR < 60%（最大损失可控）')
    add_paragraph(document, '4. 盈利概率 > 50%（盈利可能性较高）')
    add_paragraph(document, '')

    # 选择逻辑
    add_paragraph(document, ' 选择逻辑（从符合条件中优选）：', bold=True)
    add_paragraph(document, '• 优先选择最接近8%收益率的方案（保守原则）')
    add_paragraph(document, '• 同等收益率下选择报价更低的方案')
    add_paragraph(document, '• 同等情况下选择保守程度更高的方案')
    add_paragraph(document, '')

    # 从context获取历史数据情景（第六章6.2-6.5节的情景）
    comprehensive_results = context['results'].get('historical_scenarios_195', [])
    print(f" 第九章9.3.2.2节：获取到{len(comprehensive_results)}个历史数据场景")

    # 确定当前应该选择的档位
    if macro_score >= 80:
        target_tier = "高"
        tier_keywords = ["75%", "高", "upper", "high"]
    elif macro_score >= 60:
        target_tier = "中"
        tier_keywords = ["50%", "中", "middle", "mid"]
    else:
        target_tier = "低"
        tier_keywords = ["25%", "低", "lower", "low"]

    # 收集情景方案（仅筛选6.2-6.5的情景：市场指数、行业指数、行业PE、个股PE、DCF估值）
    scenario_options = []

    if comprehensive_results:
        # 从comprehensive_results提取所有情景
        for result in comprehensive_results:
            # 兼容两种数据结构：
            # 1. 嵌套结构：{'scenario': {...}, 'median_return': ..., 'profit_prob': ...}
            # 2. 扁平结构：{'name': ..., 'drift_level': ..., 'profit_prob': ..., 'median_return': ...}
            if 'scenario' in result:
                # 嵌套结构
                scenario_obj = result['scenario']
                scenario_name = scenario_obj['name']
                profit_prob = result['profit_prob']
                median_return = result['median_return']
                var_5 = result.get('var_5', result.get('var_95', 0))
                issue_price = result.get('issue_price', scenario_obj.get('issue_price', 0))

                # 只筛选6.2-6.5的情景（市场指数、行业指数、行业PE、个股PE、DCF估值）
                if not any(keyword in scenario_name for keyword in ['市场指数', '行业指数', '行业PE', '个股PE', 'DCF估值']):
                    continue

                # 根据市场环境评分筛选对应档位的场景
                # 档位信息存储在drift_level和vol_level字段中，不是在名称里
                is_target_tier = False
                drift_level = scenario_obj.get('drift_level', '')
                vol_level = scenario_obj.get('vol_level', '')

                # DCF估值特殊处理：只有1个漂移率，所以没有drift_level档位，只需检查波动率档位
                if scenario_name.startswith('DCF估值'):
                    is_target_tier = (vol_level == target_tier)
                else:
                    # 其他历史数据场景：需要漂移率和波动率档位都匹配
                    if target_tier == "高":
                        # 检查漂移率和波动率是否都为高档位
                        is_target_tier = (drift_level == "高" and vol_level == "高")
                    elif target_tier == "中":
                        # 检查漂移率和波动率是否都为中档位
                        is_target_tier = (drift_level == "中" and vol_level == "中")
                    else:  # 低档位
                        # 检查漂移率和波动率是否都为低档位
                        is_target_tier = (drift_level == "低" and vol_level == "低")

                # 如果不符合目标档位，跳过该场景
                if not is_target_tier:
                    continue

                # 数据类型转换（确保格式一致）
                median_return = median_return * 100  # 从小数转换为百分比
                # 修正：使用var_5（5%分位数）作为95% VaR，表示95%置信水平下的最大损失
                var_95 = var_5 * 100  # 优先使用var_5

                scenario_options.append({
                    'name': scenario_name,
                    'description': scenario_obj.get('description', ''),
                    'median_return': median_return,
                    'profit_prob': profit_prob,
                    'premium_rate': scenario_obj.get('premium_rate', scenario_obj.get('discount', 0)) * 100,  # 转为百分比
                    'var_95': var_95,
                    'issue_price': issue_price
                })

            elif 'name' in result and 'median_return' in result and 'profit_prob' in result:
                # 扁平结构（第六章存储的格式）
                scenario_obj = result
                scenario_name = result['name']
                profit_prob = result['profit_prob']
                median_return = result['median_return']
                var_5 = result.get('var_5', result.get('var_95', 0))
                issue_price = result.get('issue_price', 0)

                # 只筛选6.2-6.5的情景（市场指数、行业指数、行业PE、个股PE、DCF估值）
                if not any(keyword in scenario_name for keyword in ['市场指数', '行业指数', '行业PE', '个股PE', 'DCF估值']):
                    continue

                # 根据市场环境评分筛选对应档位的场景
                # 档位信息存储在drift_level和vol_level字段中，不是在名称里
                is_target_tier = False
                drift_level = scenario_obj.get('drift_level', '')
                vol_level = scenario_obj.get('vol_level', '')

                # DCF估值特殊处理：只有1个漂移率，所以没有drift_level档位，只需检查波动率档位
                if scenario_name.startswith('DCF估值'):
                    is_target_tier = (vol_level == target_tier)
                else:
                    # 其他历史数据场景：需要漂移率和波动率档位都匹配
                    if target_tier == "高":
                        # 检查漂移率和波动率是否都为高档位
                        is_target_tier = (drift_level == "高" and vol_level == "高")
                    elif target_tier == "中":
                        # 检查漂移率和波动率是否都为中档位
                        is_target_tier = (drift_level == "中" and vol_level == "中")
                    else:  # 低档位
                        # 检查漂移率和波动率是否都为低档位
                        is_target_tier = (drift_level == "低" and vol_level == "低")

                # 如果不符合目标档位，跳过该场景
                if not is_target_tier:
                    continue

                # 数据类型转换（确保格式一致）
                median_return = median_return * 100  # 从小数转换为百分比
                # 修正：使用var_5（5%分位数）作为95% VaR，表示95%置信水平下的最大损失
                var_95 = var_5 * 100  # 优先使用var_5

                scenario_options.append({
                    'name': scenario_name,
                    'description': scenario_obj.get('description', ''),
                    'median_return': median_return,
                    'profit_prob': profit_prob,
                    'premium_rate': scenario_obj.get('premium_rate', scenario_obj.get('discount', 0)) * 100,  # 转为百分比
                    'var_95': var_95,
                    'issue_price': issue_price
                })

            else:
                # 数据格式不匹配，跳过
                continue

    # 如果有情景方案，进行筛选和统计
    if scenario_options:
        # 按大场景分组（5个历史数据场景）
        scenario_categories = {
            '市场指数': [],
            '行业指数': [],
            '行业PE': [],
            '个股PE': [],
            'DCF估值': []
        }

        # 将情景按大场景分类
        for scenario in scenario_options:
            for category in scenario_categories.keys():
                if category in scenario['name']:
                    scenario_categories[category].append(scenario)
                    break

        # 对每个大场景应用基础筛选条件
        qualified_by_category = {}
        for category, scenarios in scenario_categories.items():
            qualified = []
            for scenario in scenarios:
                # 检查四个基础条件
                # 注意：premium_rate和median_return已经是百分比形式，直接比较
                condition_1 = scenario['premium_rate'] <= -5  # 溢价率 ≤ -5%
                condition_2 = scenario['median_return'] > 8  # 中位数收益率 > 8%
                condition_3 = abs(scenario['var_95']) < 60  # |95% VaR| < 60%
                condition_4 = scenario['profit_prob'] > 50  # 盈利概率 > 50%

                # 显示前几个场景的条件检查结果（调试用）
                if len(qualified) == 0 and scenario['premium_rate'] > -20:
                    print(f"{category} - 场景: {scenario['name']}")
                    print(f"  条件检查: c1={condition_1}(溢价率{scenario['premium_rate']:.1f}%), c2={condition_2}(中位数收益{scenario['median_return']:.1f}%), c3={condition_3}(VaR{scenario['var_95']:.1f}%), c4={condition_4}(盈利概率{scenario['profit_prob']:.1f}%)")

                if condition_1 and condition_2 and condition_3 and condition_4:
                    qualified.append(scenario)
            qualified_by_category[category] = qualified

        print(f"筛选结果:")
        for category, scenarios in qualified_by_category.items():
            print(f"  {category}: {len(scenarios)}个场景符合条件")

        # 统计结果
        qualified_categories = {cat: scenarios for cat, scenarios in qualified_by_category.items() if scenarios}
        total_qualified_categories = len(qualified_categories)
        total_qualified_scenarios = sum(len(scenarios) for scenarios in qualified_by_category.values())

        # 显示统计结果
        add_paragraph(document, '历史数据场景筛选统计：', bold=True)
        add_paragraph(document, f'• 大场景总数：5个（市场指数、行业指数、行业PE、个股PE、DCF估值）')
        add_paragraph(document, f'• 符合条件的大场景数：{total_qualified_categories}个')
        add_paragraph(document, f'• 符合条件的总场景数：{total_qualified_scenarios}个')
        add_paragraph(document, '')

        # 显示各场景筛选结果
        for category, scenarios in qualified_by_category.items():
            category_count = len(scenarios)

            if category_count > 0:
                add_paragraph(document, f'{category}：{category_count}个场景符合条件', bold=True)

                # 创建该类别符合条件场景表格
                category_scenarios_data = []
                for scenario in scenarios:
                    category_scenarios_data.append([
                        scenario['name'],
                        f"{scenario['premium_rate']:+.1f}%",
                        f"{scenario['median_return']:+.1f}%",
                        f"{scenario['var_95']:+.1f}%",
                        f"{scenario['profit_prob']:.1f}%"
                    ])

                add_table_data(document, ['情景方案', '溢价率', '中位数收益率', '95% VaR', '盈利概率'], category_scenarios_data)
                add_paragraph(document, '')
            else:
                add_paragraph(document, f'{category}：无场景符合条件', bold=True)

        # 投资建议判断
        add_paragraph(document, '投资建议：', bold=True)
        if total_qualified_categories >= 2:
            # 至少两个大场景符合条件，支持报价
            # 收集所有符合条件场景的溢价率，形成区间
            all_qualified = []
            for scenarios in qualified_by_category.values():
                all_qualified.extend(scenarios)

            # 计算溢价率区间
            all_premiums = [s['premium_rate'] for s in all_qualified]
            min_premium = min(all_premiums)  # 最低（最大折价）
            max_premium = max(all_premiums)  # 最高（最小折价）

            add_paragraph(document, f'• 至少{total_qualified_categories}个大场景符合条件，支持参与报价')
            add_paragraph(document, f'• 指导溢价率：{min_premium:+.2f}% 至 {max_premium:+.2f}%（综合{len(all_qualified)}个符合条件的场景）')
            add_paragraph(document, f'• 建议报价范围：溢价率在[{min_premium:+.2f}%, {max_premium:+.2f}%]区间内（折价{abs(min_premium):.1f}%至{abs(max_premium):.1f}%）')
        else:
            # 少于两个大场景符合条件，建议不参与
            add_paragraph(document, f'• 仅{total_qualified_categories}个大场景符合条件（少于2个），建议本项目不参与')
            add_paragraph(document, '• 风险提示：历史数据场景支持不足，建议谨慎考虑或放弃本次定增')

        add_paragraph(document, '')
    else:
        add_paragraph(document, '历史数据场景筛选统计：', bold=True)
        add_paragraph(document, '• 无历史数据场景可供分析')
        add_paragraph(document, '• 原因：第六章情景分析未生成或数据缺失')
        add_paragraph(document, '• 建议：请检查第六章情景分析是否正常生成')
        
    # ==================== 9.3.2.3 基于参数构造场景的筛选 ====================
    add_title(document, '9.3.2.3 基于参数构造场景的筛选', level=4)

    add_paragraph(document, '本节从585种参数构造场景（漂移率13档 × 波动率5档 × 溢价率9档）中筛选符合条件的方案，提供更全面的报价参考。')

    # 从context获取多参数构造情景（第六章6.1节的情景）
    comprehensive_results = context['results'].get('multi_param_scenarios_585', [])

    # 第一步：筛选符合基础条件的585种参数构造场景
    qualified_585_scenarios = []
    checked_count = 0

    for result in comprehensive_results:
        # 兼容新旧格式
        scenario_obj = result.get('scenario', result)

        # 只处理多参数情景（名称不带"市场指数"、"行业指数"等前缀的）
        scenario_name = scenario_obj.get('name', '')
        if any(prefix in scenario_name for prefix in ['市场指数', '行业指数', '行业PE', '个股PE', 'DCF估值']):
            continue

        checked_count += 1

        # 获取实际溢价率（百分比形式）
        premium_rate_pct = scenario_obj.get('premium_rate', scenario_obj.get('discount', 0)) * 100
        median_return_pct = result.get('median_return', 0) * 100
        var_5_raw = result.get('var_5', result.get('var_95', 0)) * 100  # 原始值，可能是负值
        var_95_pct = abs(var_5_raw)  # 绝对值，用于筛选条件
        profit_prob_pct = result.get('profit_prob', 0)

        # 检查四个基础条件
        condition_1 = premium_rate_pct <= -5  # 溢价率 ≤ -5%
        condition_2 = median_return_pct > 8  # 中位数收益率 > 8%
        condition_3 = var_95_pct < 60  # |95% VaR| < 60%
        condition_4 = profit_prob_pct > 50  # 盈利概率 > 50%

        if condition_1 and condition_2 and condition_3 and condition_4:
            qualified_585_scenarios.append({
                'name': scenario_name,
                'drift': scenario_obj['drift'],
                'volatility': scenario_obj['volatility'],
                'premium_rate': premium_rate_pct,
                'issue_price': scenario_obj.get('issue_price', 0),
                'median_return': median_return_pct,
                'var_95': var_5_raw,
                'profit_prob': profit_prob_pct
            })

    if qualified_585_scenarios:
        add_paragraph(document, f'参数构造场景筛选结果：', bold=True)
        add_paragraph(document, f'• 从585种参数构造场景中找到{len(qualified_585_scenarios)}个符合基础条件的方案')

        # 第二步：根据市场环境评分进行参数映射
        add_paragraph(document, '根据市场环境评分进行参数映射：', bold=True)
        add_paragraph(document, f'• 参考9.3.1节的市场环境评分参数映射表')
        add_paragraph(document, '')

        add_paragraph(document, '当前项目参数映射结果：', bold=True)

        # 使用前面定义的映射函数获取当前项目的参数范围
        drift_min, drift_max, vol_min, vol_max = get_drift_volatility_mapping(macro_score)

        add_paragraph(document, f'• 当前宏观评分：{macro_score}分')
        add_paragraph(document, f'• 对应漂移率范围：{drift_min*100:.0f}% - {drift_max*100:.0f}%')
        add_paragraph(document, f'• 对应波动率范围：{vol_min*100:.0f}% - {vol_max*100:.0f}%')
        add_paragraph(document, '')

        # 第三步：从所有符合基础条件的场景中筛选符合市场环境参数的场景
        environment_matched_scenarios = []
        for scenario in qualified_585_scenarios:
            drift = scenario['drift']
            volatility = scenario['volatility']

            # 检查是否在对应范围内
            if drift_min <= drift <= drift_max and vol_min <= volatility <= vol_max:
                environment_matched_scenarios.append(scenario)

        add_paragraph(document, f'环境匹配筛选结果：', bold=True)
        if environment_matched_scenarios:
            add_paragraph(document, f'• 从{len(qualified_585_scenarios)}个符合基础条件的场景中找到{len(environment_matched_scenarios)}个符合当前市场环境的方案')

            # 取符合环境场景中能接受的最大溢价率（即溢价率最高的）
            environment_matched_scenarios.sort(key=lambda x: x['premium_rate'], reverse=True)
            best_matched_scenario = environment_matched_scenarios[0]
            environment_premium_threshold = best_matched_scenario['premium_rate']

            add_paragraph(document, f'• 环境匹配阈值溢价率：{environment_premium_threshold:+.2f}%')
            add_paragraph(document, f'• 来源于：漂移率{best_matched_scenario["drift"]*100:.0f}%, 波动率{best_matched_scenario["volatility"]*100:.0f}%')
            add_paragraph(document, '')

            # 显示推荐方案详情
            add_paragraph(document, '参数构造场景推荐方案：', bold=True)
            recommended_585_data = [
                ['建议报价', f"{best_matched_scenario['issue_price']:.2f}元/股", f"溢价率{best_matched_scenario['premium_rate']:+.1f}%"],
                ['预期中位数收益率', f"{best_matched_scenario['median_return']:+.1f}%", '年化收益率'],
                ['盈利概率', f"{best_matched_scenario['profit_prob']:.1f}%", '蒙特卡洛模拟盈利概率'],
                ['95% VaR', f"{best_matched_scenario['var_95']:+.1f}%", '锁定期最大损失（负值）'],
                ['', '', ''],
                ['市场参数', f"漂移率{best_matched_scenario['drift']*100:+.0f}%, 波动率{best_matched_scenario['volatility']*100:.0f}%", '基于市场环境映射'],
                [' 溢价率', f"{best_matched_scenario['premium_rate']:+.1f}%", '≤ -5% '],
                [' 中位数收益率', f"{best_matched_scenario['median_return']:+.1f}%", '> 8% '],
                [' 95% VaR', f"{best_matched_scenario['var_95']:+.1f}%", '|VaR| < 60% '],
                [' 盈利概率', f"{best_matched_scenario['profit_prob']:.1f}%", '> 50% '],
            ]

            add_table_data(document, ['指标', '数值', '说明'], recommended_585_data)
        else:
            add_paragraph(document, f'• 前10个场景中没有完全符合当前市场环境（漂移率{drift_min*100:.0f}-{drift_max*100:.0f}%, 波动率{vol_min*100:.0f}-{vol_max*100:.0f}%）的方案')
            add_paragraph(document, '• 建议适当放宽市场环境要求或调整参数范围')
    else:
        add_paragraph(document, '参数构造场景筛选结果：', bold=True)
        add_paragraph(document, '• 从585种参数构造场景中未找到符合基础条件的方案')
        add_paragraph(document, '• 建议调整筛选条件或等待更好的市场时机')

    
    # ==================== 9.3.2.4 综合溢价率区间建议 ====================
    add_title(document, '9.3.2.4 综合溢价率区间建议', level=4)

    add_paragraph(document, '综合以上三种情景筛选结果，给出最终的溢价率区间建议。')
    add_paragraph(document, '')

    # 9.3.2.4.1 溢价率情景对照表
    add_paragraph(document, '9.3.2.4.1 溢价率情景对照表', bold=True)

    add_paragraph(document, '为明确推导过程，以下是各筛选方法中每一档溢价率对应的情景统计：')
    add_paragraph(document, '')

    # 创建综合的溢价率情景对照表
    # 使用已筛选的符合条件的历史数据场景，而不是所有场景
    if 'qualified_by_category' in locals() and qualified_by_category:
        # 按溢价率分组统计
        premium_groups = {}
        for category, scenarios in qualified_by_category.items():
            for scenario in scenarios:
                premium_rate = scenario['premium_rate']
                premium_key = f"{premium_rate:+.1f}%"

                if premium_key not in premium_groups:
                    premium_groups[premium_key] = {
                        'premium_rate': premium_rate,
                        'scenarios': [],
                        'categories': set(),
                        'count': 0,
                        'median_return': scenario['median_return'],
                        'profit_prob': scenario['profit_prob'],
                        'var_95': scenario.get('var_95', 0)
                    }

                premium_groups[premium_key]['scenarios'].append(scenario['name'])
                premium_groups[premium_key]['categories'].add(category)  # 直接使用category
                premium_groups[premium_key]['count'] += 1

        # 获取所有溢价率并排序
        sorted_premiums = sorted(premium_groups.keys(), key=lambda x: float(x.replace('%', '')), reverse=True)

        # 构建表格数据
        premium_table_data = []
        for premium_key in sorted_premiums:
            group = premium_groups[premium_key]
            scenarios_list = list(group['categories'])
            scenarios_list.sort()

            # 获取该档次第一个情景的详细信息（从已筛选的情景中）
            first_scenario = None
            for category, scenarios in qualified_by_category.items():
                for scenario in scenarios:
                    if f"{scenario['premium_rate']:+.1f}%" == premium_key:
                        first_scenario = scenario
                        break
                if first_scenario:
                    break

            # 情景特征说明
            scenario_features = []
            if first_scenario:
                if first_scenario['median_return'] > 8:
                    scenario_features.append(f"收益率{first_scenario['median_return']:.1f}%")
                if first_scenario['profit_prob'] > 50:
                    scenario_features.append(f"盈利概率{first_scenario['profit_prob']:.1f}%")
                if abs(first_scenario.get('var_95', 0)) < 60:
                    scenario_features.append(f"风险可控")

            premium_table_data.append([
                premium_key,
                f"{group['count']}个",
                ', '.join(scenarios_list),
                '; '.join(scenario_features) if scenario_features else "基础符合条件"
            ])

        premium_headers = ['溢价率档次', '符合情景数量', '情景类型', '情景特征']
        add_table_data(document, premium_headers, premium_table_data)

        add_paragraph(document, '')
        add_paragraph(document, '表格说明：')
        add_paragraph(document, '• 溢价率档次：按名义溢价率（相对MA20）从高到低排序，数值越大折价越大')
        add_paragraph(document, '• 符合情景数量：该档次中满足筛选条件的情景总数，反映历史数据支持程度')
        add_paragraph(document, '• 情景类型：涉及的5个大类（市场指数、行业指数、行业PE、个股PE、DCF估值）')
        add_paragraph(document, '• 情景特征：该档次情景的关键指标特征（收益率、盈利概率、风险水平）')
        add_paragraph(document, '')

        # 添加推导说明
        add_paragraph(document, '区间推导逻辑：', bold=True)
        add_paragraph(document, '• 最终报价区间基于上述表格中的溢价率档次分布综合确定')
        add_paragraph(document, '• 优先选择符合条件数量较多的溢价率区间，确保历史数据支持充分')
        add_paragraph(document, '• 结合反向推算和参数构造情景的约束条件，确定区间的上下限')
        add_paragraph(document, '• 综合考虑三种筛选方法（历史数据、参数构造、反向推算）的交集结果')
        add_paragraph(document, '')

    # 9.3.2.4.2 综合分析结果
    add_paragraph(document, '9.3.2.4.2 综合分析结果', bold=True)
    add_paragraph(document, '')

    # 这里需要汇总前面三个子节的阈值结果
    # 由于前面结果是在不同作用域中计算的，这里需要重新计算或存储
    # 为了简化，我们在这里重新计算关键的阈值

    # 1. 历史数据场景阈值（如果有符合条件的）
    historical_threshold = None
    if 'scenario_options' in locals():
        historical_qualified = []
        for scenario in scenario_options:
            condition_1 = scenario['premium_rate'] <= -5
            condition_2 = scenario['median_return'] > 8
            condition_3 = abs(scenario['var_95']) < 60
            condition_4 = scenario['profit_prob'] > 50

            if condition_1 and condition_2 and condition_3 and condition_4:
                historical_qualified.append(scenario)

        if historical_qualified and len(qualified_by_category) >= 2:
            historical_qualified.sort(key=lambda x: x['premium_rate'], reverse=True)
            historical_threshold = historical_qualified[0]['premium_rate']

    # 2. 参数构造场景阈值（如果有符合条件的）
    param_threshold = None
    param_threshold_actual = None
    if 'environment_matched_scenarios' in locals() and environment_matched_scenarios:
        param_nominal_premium = environment_matched_scenarios[0]['premium_rate']  # 名义溢价率（相对MA20）
        # 基于名义溢价率计算发行价：发行价 = MA20 × (1 + 名义溢价率)
        issue_price_from_nominal = ma20_price * (1 + param_nominal_premium / 100)
        # 计算实际溢价率（相对当前价格）：实际溢价率 = 发行价 / 当前价格 - 1
        param_threshold_actual = (issue_price_from_nominal / current_price_eval - 1) * 100
        # 名义溢价率（保持不变）
        param_threshold = param_nominal_premium

    # 3. 反向推算阈值（使用名义溢价率相对MA20）
    reverse_nominal_premium = premium_to_ma20_adjusted  # 从9.3.2.1节获取名义溢价率（相对MA20）
    # 基于名义溢价率计算发行价：发行价 = MA20 × (1 + 名义溢价率)
    issue_price_from_nominal = ma20_price * (1 + reverse_nominal_premium / 100)
    # 计算实际溢价率（相对当前价格）：实际溢价率 = 发行价 / 当前价格 - 1
    reverse_threshold_actual = (issue_price_from_nominal / current_price_eval - 1) * 100
    # 名义溢价率（保持不变）
    reverse_threshold = reverse_nominal_premium

    # 汇总结果（统一使用名义溢价率）
    thresholds = []
    if historical_threshold is not None:
        thresholds.append(('历史数据场景', historical_threshold))
    if param_threshold is not None:
        thresholds.append(('参数构造场景', param_threshold))
    if reverse_threshold is not None:
        thresholds.append(('反向推算', reverse_threshold))

    if len(thresholds) >= 2:
        # 统一使用名义溢价率（相对MA20）
        all_premiums = [t[1] for t in thresholds]

        min_premium = min(all_premiums)
        max_premium = max(all_premiums)
        premium_range = max_premium - min_premium

        add_paragraph(document, f'• 有效阈值数量：{len(thresholds)}个')
        for name, threshold in thresholds:
            add_paragraph(document, f'• {name}阈值：{threshold:+.2f}%（名义溢价率，相对MA20）')

        add_paragraph(document, '')
        add_paragraph(document, '最终溢价率区间建议：', bold=True)
        add_paragraph(document, f'• 名义溢价率区间（相对MA20）：[{min_premium:+.2f}%, {max_premium:+.2f}%]')
        add_paragraph(document, f'• 区间宽度：{premium_range:.2f}%')

        add_paragraph(document, '')
        add_paragraph(document, '最终建议：', bold=True)
        add_paragraph(document, f'• 名义溢价率上限：≤{max_premium:.2f}%（相对MA20，硬约束）')
        add_paragraph(document, f'• 名义溢价率下限：{min_premium:+.2f}%（相对MA20，最优方案）')
        add_paragraph(document, f'• 区间宽度：{premium_range:.2f}%')
    else:
        add_paragraph(document, '• 有效阈值数量不足（少于2个）')
        add_paragraph(document, '• 无法形成可靠的溢价率区间，建议谨慎参与本次定增')
        add_paragraph(document, '• 或参考单一可用的阈值进行报价决策')

        # ==================== 9.4 主要风险提示 ====================
    add_title(document, '9.4 主要风险提示', level=2)
    add_paragraph(document, '基于多维度风险分析，提示以下主要风险（详见前文各章节详细分析）：')
    add_paragraph(document, '')

    # 1. 市场风险
    add_paragraph(document, '1. 市场风险')
    add_paragraph(document, f'   • 波动率风险：当前120日窗口年化波动率为{mc_volatility_120d*100:.1f}%，市场波动可能导致实际收益偏离预期')
    add_paragraph(document, f'   • 趋势风险：当前120日窗口年化漂移率为{mc_drift_120d*100:+.2f}%，{"上升趋势" if mc_drift_120d > 0 else "下降趋势" if mc_drift_120d < 0 else "震荡趋势"}可能影响解禁时收益')

    # 2. 流动性风险
    add_paragraph(document, '2. 流动性风险')
    add_paragraph(document, f'   • 锁定期风险：{project_params["lockup_period"]}个月锁定期内无法交易，需承担期间价格波动')
    add_paragraph(document, '   • 解禁冲击：解禁后可能面临抛压，导致实际变现价格低于理论价格')

    # 3. VaR在险价值风险
    add_paragraph(document, '3. VaR在险价值风险')
    # 使用基于期收益率（锁定期收益率）的VaR，而不是年化VaR
    # 从context中获取var_results，然后获取120日窗口的var_95
    var_results = context.get('results', {}).get('var_results', {})
    var_95_period = var_results.get('120日', {}).get('var_95', 0.5)
    var_95_display = var_95_period * 100  # 转换为百分比
    add_paragraph(document, f'   • 120日窗口：95%置信水平下最大可能亏损{var_95_display:.1f}%')
    add_paragraph(document, '   • 尾部风险：历史数据显示，小概率极端事件（黑天鹅）可能导致损失超过VaR预测值')

    # 4. 估值风险
    add_paragraph(document, '4. 估值风险')
    add_paragraph(document, f'   • DCF估值风险：DCF内在价值{intrinsic_value:.2f}元/股基于多个假设，实际业绩可能偏离预测')
    add_paragraph(document, '   • 相对估值风险：PE/PS/PB相对估值基于行业平均水平，行业景气度变化可能导致估值体系重构')

    # 5. 发行定价风险
    add_paragraph(document, '5. 发行定价风险')
    add_paragraph(document, '   • 溢价风险：如发行价过高，影响安全边际和整体抗风险能力。')
    add_paragraph(document, '   • 溢价发行无安全边际，需重点关注公司成长性')
    add_paragraph(document, '   • 定价偏离：若发行价显著高于盈亏平衡价格，将大幅降低盈利概率和预期收益')

    # 6. 其他风险
    add_paragraph(document, '6. 其他风险')
    add_paragraph(document, '   • 行业政策风险：需关注行业监管政策变化')
    add_paragraph(document, '   • 业绩波动风险：需关注公司业绩预告、审计报告等')
    add_paragraph(document, '   • 竞争格局风险：行业竞争加剧可能影响盈利能力')

    # ==================== 9.5 免责声明 ====================
    add_title(document, '9.5 免责声明', level=2)

    disclaimer = f'''
    本报告数据取至tushare平台，报告系半自动化生成，仅供参考使用，不构成投资建议。

    1. 本报告基于历史数据和公开信息进行分析，不能保证tushare数据不会出错；
    2. 市场有风险，投资需谨慎。本报告中的任何分析观点不代表未来表现；
    3. 本报告提到的任何证券或投资标的，仅为分析示例，不构成推荐；
    4. 本报告的知识产权归分析团队所有，未经许可不得转载或使用。

    报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    '''

    add_paragraph(document, disclaimer, font_size=9)

    add_section_break(document)

    return context
