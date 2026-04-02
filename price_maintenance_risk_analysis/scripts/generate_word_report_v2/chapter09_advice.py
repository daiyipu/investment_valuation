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


def calculate_macro_environment_assessment(market_data, document):
    """
    计算宏观环境评估得分

    评估维度：
    1. 货币政策（15%权重）：紧缩(40分)/稳健(60分)/适度宽松(80分)/宽松(100分)
    2. 财政政策（15%权重）：紧缩(50分)/稳健(70分)/积极(90分)
    3. 行业发展周期（30%权重）：成长(100分)/成熟(80分)/幼稚(60分)/衰退(40分)
    4. 二级市场活跃度（40%权重）：基于历史分位数

    Args:
        market_data: 市场数据字典
        document: Word文档对象（用于添加表格）

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

    # 如果无法判断，给予中性分数50分
    if industry_pe == 0 or current_pe == 0:
        industry_cycle_score = 50
        industry_cycle_status = "待判断"

    # ==================== 4. 二级市场活跃度评估（40%权重）====================
    # 基于历史换手率分位数（最近5年）
    # 这里简化处理：基于当前波动率和成交量

    # 获取历史换手率分位数（如果有）
    # 如果没有，使用价格相对MA20的位置作为代理指标
    current_price = market_data.get('current_price', 20)
    ma20 = market_data.get('ma_20', current_price)
    ma120 = market_data.get('ma_120', current_price)

    # 价格相对MA20和MA120的位置
    price_to_ma20 = current_price / ma20 if ma20 > 0 else 1.0
    price_to_ma120 = current_price / ma120 if ma120 > 0 else 1.0

    # 基于价格位置评估市场活跃度
    if price_to_ma120 > 1.2:
        market_activity_score = 100  # 强势
        market_activity_status = "活跃"
    elif price_to_ma120 > 1.05:
        market_activity_score = 85.7  # 偏强势（历史分位数85.7%）
        market_activity_status = "活跃"
    elif price_to_ma120 > 0.95:
        market_activity_score = 60  # 中性
        market_activity_status = "一般"
    elif price_to_ma120 > 0.85:
        market_activity_score = 40  # 偏弱
        market_activity_status = "低迷"
    else:
        market_activity_score = 20  # 弱势
        market_activity_status = "极度低迷"

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
            industry_cycle_status,
            f'{industry_cycle_score:.0f}',
            '30%',
            f'{weighted_scores["industry_cycle"]:.1f}',
            '成长(100分)/成熟(80分)/幼稚(60分)/衰退(40分) - 请人工判断填写'
        ],
        [
            '二级市场活跃度',
            market_activity_status,
            f'{market_activity_score:.1f}',
            '40%',
            f'{weighted_scores["market_activity"]:.1f}',
            f'历史分位数>{int(market_activity_score)}%（5年，当前{price_to_ma20*100:.1f}%）'
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
    add_paragraph(document, '')

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
    生成第九章：风控建议与风险提示

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
    add_paragraph(document, '')

    # ==================== 9.2 盈亏平衡分析 ====================
    add_title(document, '9.2 盈亏平衡分析', level=2)
    add_paragraph(document, '本节通过盈亏平衡分析，评估在不同目标收益率和不同退出周期下的安全边际。')
    add_paragraph(document, '')

    # ==================== 9.2.1 基于当前锁定期的盈亏平衡分析 ====================
    add_title(document, '9.2.1 基于当前锁定期的盈亏平衡分析', level=3)

    add_paragraph(document, '本节基于当前锁定期，分析不同目标收益率下的盈亏平衡价格。')
    add_paragraph(document, '')

    # 计算不同收益率下的盈亏平衡价格
    import numpy as np
    target_returns = np.linspace(0.05, 0.50, 10)
    break_even_prices = []
    issue_price = project_params['issue_price']
    issue_price_eval = issue_price  # 添加此变量供后续使用
    lockup_years = project_params['lockup_period'] / 12
    current_price_eval = project_params['current_price']

    for target_return in target_returns:
        be_price = issue_price * (1 + target_return * lockup_years)
        break_even_prices.append(be_price)

    # 生成表格数据
    be_data = []
    for ret, price in zip(target_returns[::2], break_even_prices[::2]):
        distance = (current_price_eval - price) / current_price_eval * 100
        status = "" if distance > 0 else ""
        be_data.append([f'{ret*100:.0f}%', f'{price:.2f}', f'{distance:+.1f}%', status])

    add_table_data(document, ['期望年化收益率', '盈亏平衡价格(元)', '距离当前价', '状态'], be_data)

    add_paragraph(document, '')
    add_paragraph(document, '盈亏平衡分析结论：')
    add_paragraph(document, f'• 当前价格: {current_price_eval:.2f} 元')
    add_paragraph(document, f'• 发行价格: {issue_price:.2f} 元')
    add_paragraph(document, f'• 锁定期: {project_params["lockup_period"]}个月（{lockup_years:.2f}年）')
    add_paragraph(document, f'• 20%年化收益率要求下盈亏平衡价: {break_even_prices[3]:.2f} 元')

    be_20 = break_even_prices[3]
    if current_price_eval > be_20:
        margin = (current_price_eval - be_20) / current_price_eval * 100
        add_paragraph(document, f'•  当前价格高于20%收益率盈亏平衡价{margin:.1f}%，具有较好安全边际')
    else:
        gap = (be_20 - current_price_eval) / current_price_eval * 100
        add_paragraph(document, f'•  当前价格低于20%收益率盈亏平衡价{gap:.1f}%，安全边际不足')

    add_paragraph(document, '')

    # ==================== 9.2.2 考虑退出周期的盈亏平衡分析 ====================
    add_title(document, '9.2.2 考虑退出周期的盈亏平衡分析', level=3)

    add_paragraph(document, '本节考虑不同的退出周期（锁定期），计算在不同周期下的盈亏平衡价格。')
    add_paragraph(document, '分析考虑资金的时间成本，年化资金成本统一按4%计算。')
    add_paragraph(document, '')

    # 期望收益率和资金成本
    target_return = 0.08  # 期望收益率年化8%
    cost_of_capital = 0.04  # 资金成本年化4%

    add_paragraph(document, ' 计算参数：', bold=True)
    add_paragraph(document, f'• 期望收益率：{target_return*100:.0f}%（年化）')
    add_paragraph(document, f'• 资金成本：{cost_of_capital*100:.0f}%（年化）')
    add_paragraph(document, f'• 净收益率：{(target_return - cost_of_capital)*100:.0f}%（年化）')
    add_paragraph(document, f'• 当前发行价：{issue_price:.2f} 元/股')
    add_paragraph(document, f'• 当前价格：{current_price_eval:.2f} 元/股')
    add_paragraph(document, '')

    # 按照3个月一期，从6个月到24个月
    lockup_periods = [6, 9, 12, 15, 18, 21, 24]  # 单位：月
    period_analysis = []

    for months in lockup_periods:
        years = months / 12

        # 计算考虑资金成本后的盈亏平衡价
        # 公式：盈亏平衡价 = 发行价 × (1 + (期望收益率 + 资金成本) × 年数)
        # 这里需要考虑：期望收益率是投资回报要求，资金成本是融资成本
        # 所以实际需要的收益率是两者之和

        # 方案1：考虑资金成本和时间价值
        be_price_with_cost = issue_price * (1 + (target_return + cost_of_capital) * years)

        # 方案2：传统方法（仅考虑期望年化收益率的时间价值）
        # 传统方法下，盈亏平衡价应该考虑锁定期内的期望收益增长
        be_price_traditional = issue_price * (1 + target_return * years)

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
        ['退出周期', '锁定期年数', '盈亏平衡价(考虑资金成本)', '盈亏平衡价(传统)', '距离当前价(考虑成本)', '距离当前价(传统)', '评估'],
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
            f'{item["years"]:.2f}年',
            f'{item["be_price_with_cost"]:.2f}元',
            f'{item["be_price_traditional"]:.2f}元',
            f'{distance_with_cost:+.1f}%',
            f'{distance_traditional:+.1f}%',
            eval_text
        ])

    add_table_data(document, period_data[0], period_data[1:])

    add_paragraph(document, '')
    add_paragraph(document, ' 退出周期分析结论：', bold=True)
    add_paragraph(document, '说明：')
    add_paragraph(document, f'• 考虑资金成本方法：盈亏平衡价 = 发行价 × (1 + ({target_return*100:.0f}% + {cost_of_capital*100:.0f}%) × 年数)')
    add_paragraph(document, f'• 传统方法：盈亏平衡价 = 发行价 × (1 + {target_return*100:.0f}% × 年数)（考虑期望收益率的时间价值）')
    add_paragraph(document, f'• 资金成本{cost_of_capital*100:.0f}%已计入，反映资金的时间价值')
    add_paragraph(document, '')

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

    add_paragraph(document, '')
    add_paragraph(document, ' 投资建议：', bold=True)
    add_paragraph(document, '• 退出周期越长，资金成本越高，要求的盈亏平衡价也越高')
    add_paragraph(document, '• 建议选择退出周期较短（6-12个月）的项目，以降低资金成本压力')
    add_paragraph(document, '• 如果当前价格低于考虑资金成本的盈亏平衡价，建议要求更高的折价率')
    add_paragraph(document, '')

    # 生成并插入盈亏平衡价格敏感性图表
    break_even_chart_path = os.path.join(IMAGES_DIR, '09_break_even_analysis.png')
    generate_break_even_chart(issue_price, current_price_eval, project_params['lockup_period'], break_even_chart_path)
    add_paragraph(document, '图表 9.1: 盈亏平衡价格敏感性曲线')
    add_image(document, break_even_chart_path, width=Inches(6))

    add_paragraph(document, '')

    # ==================== 9.3 报价方案建议 ====================
    add_title(document, '9.3 报价方案建议', level=2)
    add_paragraph(document, '本节提供不同目标收益率下的报价建议，帮助投资者根据风险偏好选择合适的报价方案。')
    add_paragraph(document, '')

    # ==================== 宏观环境评估 ====================
    add_paragraph(document, '【宏观环境评估】', bold=True)
    add_paragraph(document, '在进行报价方案建议之前，首先评估当前宏观环境，为情景选择提供参考依据。')
    add_paragraph(document, '')

    # 计算宏观环境评分
    macro_assessment = calculate_macro_environment_assessment(market_data, document)

    # 保存宏观环境评估结果到context，供后续情景选择使用
    context['macro_environment_assessment'] = macro_assessment

    # 显示宏观环境评估结论
    total_score = macro_assessment['total_score']
    assessment_level = macro_assessment['assessment_level']

    add_paragraph(document, f'综合评分：{total_score:.1f}分 | 评估等级：{assessment_level}')
    add_paragraph(document, '')

    # 根据评估等级给出情景选择建议
    if total_score >= 90:
        add_paragraph(document, f'✓ 宏观环境"积极"（{total_score:.1f}分）：建议选择漂移率≥+10%的情景，溢价率可控制在0%至-5%')
    elif total_score >= 80:
        add_paragraph(document, f'✓ 宏观环境"适度"（{total_score:.1f}分）：建议选择漂移率0%至+10%的情景，溢价率控制在-5%至-10%')
    elif total_score >= 70:
        add_paragraph(document, f'✓ 宏观环境"稳健"（{total_score:.1f}分）：建议选择漂移率-5%至+5%的情景，溢价率控制在-10%至-15%')
    elif total_score >= 60:
        add_paragraph(document, f'⚠ 宏观环境"偏悲观"（{total_score:.1f}分）：建议选择漂移率≤0%的情景，溢价率控制在-15%至-20%')
    else:
        add_paragraph(document, f'✗ 宏观环境"悲观"（{total_score:.1f}分）：建议选择漂移率≤-10%的情景，溢价率要求≤-20%或谨慎参与')

    add_paragraph(document, '')
    add_paragraph(document, '注：上述溢价率建议基于宏观环境评估，具体报价还需结合项目基本面和市场情况综合判断。')
    add_paragraph(document, '')
    add_paragraph(document, '')

    # 9.3.1 宏观环境评估
    add_title(document, '9.3.1 宏观环境评估', level=3)
    add_paragraph(document, '在制定报价方案前，先评估当前的宏观环境，包括货币政策与财政政策、行业发展周期、二级市场活跃度三个维度。')
    add_paragraph(document, '')

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

    # 宏观环境评估表格
    add_paragraph(document, ' 宏观环境三维度评估：', bold=True)
    add_paragraph(document, '')

    # 货币政策和财政政策（固定值）
    monetary_policy = '适度宽松'
    monetary_policy_score = 80
    monetary_policy_weight = 0.15  # 15%
    monetary_policy_weighted = monetary_policy_score * monetary_policy_weight

    fiscal_policy = '积极'
    fiscal_policy_score = 90
    fiscal_policy_weight = 0.15  # 15%
    fiscal_policy_weighted = fiscal_policy_score * fiscal_policy_weight

    macro_env_data = [
        ['评估维度', '当前状态', '得分', '权重', '加权得分', '说明'],
        # 1. 货币政策（权重15%）
        ['货币政策', monetary_policy, f'{monetary_policy_score}', '15%', f'{monetary_policy_weighted:.1f}',
         '紧缩(40分)/稳健(60分)/适度宽松(80分)/宽松(100分)'],
        # 2. 财政政策（权重15%）
        ['财政政策', fiscal_policy, f'{fiscal_policy_score}', '15%', f'{fiscal_policy_weighted:.1f}',
         '紧缩(50分)/稳健(70分)/积极(90分)'],
        # 3. 行业发展周期（权重30%，需人工填写）
        ['行业发展周期', '__________', '_____', '30%', '_____', '成长(100分)/成熟(80分)/幼稚(60分)/衰退(40分) - 请人工判断填写'],
        # 4. 二级市场活跃度（权重40%，自动计算）
        ['二级市场活跃度', market_activity_level, f'{market_activity_score}', '40%', f'{market_activity_score * 0.4:.1f}',
         f'{market_activity_desc}，基于最近5年历史换手率计算'],
    ]

    # 分离表头和数据
    macro_env_headers = macro_env_data[0]
    macro_env_table_data = macro_env_data[1:]
    add_table_data(document, macro_env_headers, macro_env_table_data)

    # 评估说明
    add_paragraph(document, '')
    add_paragraph(document, '评估说明：', bold=True)
    add_paragraph(document, '• 评估周期：基于最近5年历史数据计算分位数')
    add_paragraph(document, '• 权重分配：货币政策(15%) + 财政政策(15%) + 行业发展周期(30%) + 二级市场活跃度(40%)')
    add_paragraph(document, '• 得分标准：')
    add_paragraph(document, '  - 货币政策：紧缩(40分)、稳健(60分)、适度宽松(80分)、宽松(100分)')
    add_paragraph(document, '  - 财政政策：紧缩(50分)、稳健(70分)、积极(90分)')
    add_paragraph(document, '  - 行业发展周期：成长(100分)、成熟(80分)、幼稚(60分)、衰退(40分)')
    add_paragraph(document, '  - 二级市场活跃度：根据历史分位数自动计算（20-100分）')
    add_paragraph(document, '')

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
            add_paragraph(document, '')

    add_paragraph(document, '')

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
    add_paragraph(document, '')

    # 9.3.2 报价方案建议
    add_title(document, '9.3.2 报价方案建议', level=3)
    add_paragraph(document, '本节基于历史收益率和情景分析，提供多角度的报价建议。')
    add_paragraph(document, '')

    add_title(document, '9.3.2.1 反向推算最高报价', level=4)

    # 计算参数
    target_return = 0.08  # 目标收益率8%
    lockup_years = project_params['lockup_period'] / 12
    volatility_120d = market_data.get('volatility_120d', market_data.get('volatility', 0.30))
    current_price_eval = project_params['current_price']
    ma20_price = market_data.get('ma_20', current_price_eval)

    # 方法1：简单反推（保守）- 不考虑波动率
    max_issue_price_simple = current_price_eval / (1 + target_return * lockup_years)
    premium_to_current_simple = (max_issue_price_simple - current_price_eval) / current_price_eval * 100
    premium_to_ma20_simple = (max_issue_price_simple - ma20_price) / ma20_price * 100

    # 方法2：考虑波动率的反推（使用安全边际）
    safety_margin = 0.5 * volatility_120d ** 2
    adjusted_return = target_return + safety_margin
    max_issue_price_adjusted = current_price_eval / (1 + adjusted_return * lockup_years)
    premium_to_current_adjusted = (max_issue_price_adjusted - current_price_eval) / current_price_eval * 100
    premium_to_ma20_adjusted = (max_issue_price_adjusted - ma20_price) / ma20_price * 100

    # 执行逻辑说明
    add_paragraph(document, '📐 执行逻辑说明：', bold=True)
    add_paragraph(document, '基于目标收益率8%，使用以下逻辑反推最高可接受溢价率：')
    add_paragraph(document, '')
    add_paragraph(document, f'1. 设定目标收益率：8%（年化）')
    add_paragraph(document, f'2. 使用120日历史波动率：{volatility_120d*100:.2f}%')
    add_paragraph(document, f'3. 锁定期：{project_params["lockup_period"]}个月（{lockup_years:.2f}年）')
    add_paragraph(document, f'4. 参考价格：当前价格{current_price_eval:.2f}元，MA20价格{ma20_price:.2f}元')
    add_paragraph(document, '')

    add_paragraph(document, '计算方法：', bold=True)
    add_paragraph(document, '• 方法1（保守）：简单反推，不考虑波动率')
    add_paragraph(document, '  公式：最高发行价 = 当前价格 ÷ (1 + 目标收益率 × 锁定期年数)')
    add_paragraph(document, f'  最高发行价 = {current_price_eval:.2f} ÷ (1 + 8% × {lockup_years:.2f}) = {max_issue_price_simple:.2f}元')
    add_paragraph(document, '')

    add_paragraph(document, '• 方法2（更保守）：考虑波动率安全边际')
    add_paragraph(document, f'  安全边际 = 0.5 × 波动率² = 0.5 × ({volatility_120d*100:.2f}%)² = {safety_margin*100:.2f}%')
    add_paragraph(document, f'  调整后收益率 = 目标收益率 + 安全边际 = 8% + {safety_margin*100:.2f}% = {adjusted_return*100:.2f}%')
    add_paragraph(document, f'  最高发行价 = {current_price_eval:.2f} ÷ (1 + {adjusted_return*100:.2f}% × {lockup_years:.2f}) = {max_issue_price_adjusted:.2f}元')
    add_paragraph(document, '')

    # 推荐方案对比表
    add_paragraph(document, '推荐方案对比：', bold=True)
    add_paragraph(document, '')

    pricing_options = [
        ['计算方法', '最高发行价', '溢价率（相对当前价）', '溢价率（相对MA20）', '特点'],
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
    recommended_premium = premium_to_ma20_adjusted
    recommended_price = max_issue_price_adjusted

    add_paragraph(document, '• 推荐使用方法2（考虑波动率安全边际）')
    add_paragraph(document, f'• 建议最高报价：不高于{recommended_price:.2f}元/股')
    add_paragraph(document, f'• 建议最高溢价率（相对MA20）：{recommended_premium:+.2f}%')
    add_paragraph(document, '• 该溢价率已考虑120日历史波动率的安全边际')
    add_paragraph(document, '')

    # 投资建议
    add_paragraph(document, ' 投资建议：', bold=True)

    if recommended_premium <= -5:
        add_paragraph(document, f'•  建议报价 - 溢价率{recommended_premium:+.2f}%提供{abs(recommended_premium):.2f}%的安全边际，符合保守投资原则')
    elif recommended_premium <= 0:
        add_paragraph(document, f'•  谨慎报价 - 溢价率{recommended_premium:+.2f}%安全边际有限，建议折价至少5%')
    else:
        add_paragraph(document, f'•  不建议报价 - 溢价率{recommended_premium:+.2f}%为溢价发行，缺乏安全边际')
        add_paragraph(document, '  建议：等待市场时机或选择其他投资机会')

    add_paragraph(document, '')
    add_paragraph(document, ' 注意事项：')
    add_paragraph(document, '• 本计算基于目标收益率8%，实际收益可能因市场波动而偏离')
    add_paragraph(document, '• 建议结合其他分析方法（如DCF估值、相对估值）综合判断')
    add_paragraph(document, '• 如果实际发行价高于推荐报价，建议谨慎参与或放弃')
    add_paragraph(document, '')

    # ==================== 9.3.2.2 基于情景分析的筛选结果 ====================
    add_title(document, '9.3.2.2 基于情景分析的筛选结果', level=4)

    add_paragraph(document, '本节基于第六章情景分析的780种情景，通过严格筛选条件，从风险收益平衡角度推荐最优报价方案。')
    add_paragraph(document, '')
    add_paragraph(document, '详细的情景数据请参见报告附件《情景数据表》。')
    add_paragraph(document, '')

    # 筛选条件
    add_paragraph(document, '📋 筛选范围说明：', bold=True)
    add_paragraph(document, '情景筛选分为两部分：')
    add_paragraph(document, '• 基础情景分析（585种）：漂移率13档 × 波动率5档 × 溢价率9档')
    add_paragraph(document, '• 多维情景分析（195种）：市场指数45种 + 行业指数45种 + 行业PE45种 + 个股PE45种 + DCF估值15种')
    add_paragraph(document, '')
    add_paragraph(document, '📋 筛选条件（必须全部满足）：', bold=True)
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

    # 从context获取all_scenarios（第六章所有情景分析结果）
    comprehensive_results = context['results'].get('all_scenarios', [])

    # 收集所有情景方案
    scenario_options = []

    if comprehensive_results:
        # 从comprehensive_results提取所有情景
        for result in comprehensive_results:
            if 'scenario' in result and 'median_return' in result and 'profit_prob' in result:
                scenario_obj = result['scenario']
                scenario_name = scenario_obj['name']
                median_return = result['median_return'] * 100  # 从小数转换为百分比
                profit_prob = result['profit_prob']  # 已经是百分比
                # 修正：使用var_5（5%分位数）作为95% VaR，表示95%置信水平下的最大损失
                var_95 = result.get('var_5', result.get('var_95', 0)) * 100  # 优先使用var_5

                # 兼容不同格式的issue_price
                issue_price = result.get('issue_price', scenario_obj.get('issue_price', 0))

                scenario_options.append({
                    'name': scenario_name,
                    'description': scenario_obj.get('description', ''),
                    'median_return': median_return,
                    'profit_prob': profit_prob,
                    'premium_rate': scenario_obj.get('premium_rate', scenario_obj.get('discount', 0)) * 100,  # 转为百分比
                    'var_95': var_95,
                    'issue_price': issue_price
                })

    # 如果有情景方案，显示筛选结果
    if scenario_options:
        # 筛选符合条件的方案
        qualified_scenarios = []
        for scenario in scenario_options:
            # 检查四个条件
            condition_1 = scenario['premium_rate'] <= -5  # 溢价率 ≤ -5%
            condition_2 = scenario['median_return'] > 8  # 中位数收益率 > 8%
            # 修正：95% VaR应该是负值（损失），使用绝对值比较，表示最大损失不超过60%
            condition_3 = abs(scenario['var_95']) < 60  # |95% VaR| < 60%
            condition_4 = scenario['profit_prob'] > 50  # 盈利概率 > 50%

            if condition_1 and condition_2 and condition_3 and condition_4:
                qualified_scenarios.append(scenario)

        # 显示筛选结果
        add_paragraph(document, '所有情景方案筛选结果：', bold=True)
        add_paragraph(document, '')

        # 创建所有方案表格（与6.6节格式一致）
        all_scenarios_data = []
        for scenario in scenario_options:
            # 检查是否满足条件
            condition_1 = '' if scenario['premium_rate'] <= -5 else ''
            condition_2 = '' if scenario['median_return'] > 8 else ''
            # 修正：95% VaR使用绝对值比较
            condition_3 = '' if abs(scenario['var_95']) < 60 else ''
            condition_4 = '' if scenario['profit_prob'] > 50 else ''
            all_qualified = condition_1 == '' and condition_2 == '' and condition_3 == '' and condition_4 == ''

            all_scenarios_data.append([
                scenario['name'],
                f"{scenario['premium_rate']:+.1f}%",  # 与6.6节一致：整数格式
                f"{scenario['median_return']:+.1f}%",  # 与6.6节一致：保留1位小数
                f"{scenario['var_95']:+.1f}%",  # VaR显示符号
                f"{scenario['profit_prob']:.1f}%",
                f"{condition_1} {condition_2} {condition_3} {condition_4}",
                '符合' if all_qualified else '不符合'
            ])

        add_table_data(document, ['情景方案', '溢价率', '中位数收益率', '95% VaR', '盈利概率', '条件检查', '筛选结果'], all_scenarios_data)

        add_paragraph(document, '')

        # 推荐方案
        add_paragraph(document, '推荐方案：', bold=True)
        add_paragraph(document, '')

        if qualified_scenarios:
            # 选择最接近8%收益率的方案（保守原则）
            qualified_scenarios.sort(key=lambda x: abs(x['median_return'] - 8))

            # 如果收益率距离相同，选择报价更低的
            best_scenario = qualified_scenarios[0]
            close_to_8 = [s for s in qualified_scenarios if abs(s['median_return'] - 8) < 0.01]
            if len(close_to_8) > 1:
                close_to_8.sort(key=lambda x: x['issue_price'])
                best_scenario = close_to_8[0]

            # 显示推荐方案详情
            recommended_data = [
                ['推荐方案', best_scenario['name'], '基于筛选条件和选择逻辑'],
                ['建议报价', f"{best_scenario['issue_price']:.2f}元/股", f"溢价率{best_scenario['premium_rate']:+.2f}%"],
                ['预期中位数收益率', f"{best_scenario['median_return']:+.2f}%", '年化收益率'],
                ['盈利概率', f"{best_scenario['profit_prob']:.1f}%", '蒙特卡洛模拟盈利概率'],
                ['95% VaR', f"{best_scenario['var_95']:+.2f}%", '锁定期最大损失（负值表示损失）'],
                ['', '', ''],
                [' 溢价率', f"{best_scenario['premium_rate']:+.2f}%", '≤ -5% ✓'],
                [' 中位数收益率', f"{best_scenario['median_return']:+.2f}%", '> 8% ✓'],
                [' 95% VaR', f"{best_scenario['var_95']:+.2f}%", '|VaR| < 60% ✓'],
                [' 盈利概率', f"{best_scenario['profit_prob']:.1f}%", '> 50% ✓'],
            ]

            add_table_data(document, ['指标', '数值', '说明'], recommended_data)

            # 推荐理由
            add_paragraph(document, '')
            add_paragraph(document, '推荐理由：', bold=True)
            add_paragraph(document, f'• 该方案（{best_scenario["name"]}）满足所有4个筛选条件')
            add_paragraph(document, f'• 中位数收益率{best_scenario["median_return"]:.2f}%最接近8%目标，符合保守原则')
            add_paragraph(document, f'• 溢价率{best_scenario["premium_rate"]:+.2f}%提供{abs(best_scenario["premium_rate"]):.2f}%的安全边际')
            add_paragraph(document, f'• 盈利概率{best_scenario["profit_prob"]:.1f}%，超过半数场景盈利')
            # 修正：VaR是负值表示损失，需要明确说明
            var_value = best_scenario["var_95"]
            var_description = f"{abs(var_value):.2f}%的损失" if var_value < 0 else f"{var_value:.2f}%的收益"
            add_paragraph(document, f'• 95% VaR为{var_value:+.2f}%（最大可能{var_description}），风险相对可控')

            # 投资建议
            add_paragraph(document, '')
            add_paragraph(document, ' 投资建议：', bold=True)

            if best_scenario['median_return'] >= 15:
                add_paragraph(document, f'•  强烈推荐 - 该方案预期收益率{best_scenario["median_return"]:.2f}%较高，风险可控')
            elif best_scenario['median_return'] >= 10:
                add_paragraph(document, f'•  推荐 - 该方案预期收益率{best_scenario["median_return"]:.2f}%达标，风险收益平衡')
            else:
                add_paragraph(document, f'•  谨慎推荐 - 该方案预期收益率{best_scenario["median_return"]:.2f}%勉强达标，需密切关注风险')

            if best_scenario['profit_prob'] >= 70:
                add_paragraph(document, f'• 盈利概率{best_scenario["profit_prob"]:.1f}%较高，投资安全性较好')
            elif best_scenario['profit_prob'] >= 60:
                add_paragraph(document, f'• 盈利概率{best_scenario["profit_prob"]:.1f}%中等，需评估风险承受能力')
            else:
                add_paragraph(document, f'•  盈利概率{best_scenario["profit_prob"]:.1f}%偏低，建议提高安全边际要求')

        else:
            # 没有符合条件的方案
            add_paragraph(document, ' 情景分析中未找到完全符合所有4个筛选条件的方案')
            add_paragraph(document, '')
            add_paragraph(document, '建议：')
            add_paragraph(document, '• 适当降低预期收益率要求（如从8%降至5%）')
            add_paragraph(document, '• 要求更高的折价率（如-10%或更低）')
            add_paragraph(document, '• 谨整报价策略，等待更好的市场时机')

        add_paragraph(document, '')
        add_paragraph(document, '')

    # ==================== 9.3.2.3 基于585种细化情景的筛选 ====================
    add_title(document, '9.3.2.3 基于585种细化情景的筛选', level=4)

    add_paragraph(document, '本节从585种细化情景组合中筛选符合条件的方案，提供更全面的报价参考。')
    add_paragraph(document, '详细的585种情景数据请参见报告附件《情景数据表》。')
    add_paragraph(document, '')

    # 从context获取all_scenarios（第六章所有情景分析结果）
    comprehensive_results = context['results'].get('all_scenarios', [])

    # 筛选符合条件的585种情景
    qualified_585_scenarios = []
    for result in comprehensive_results:
        # 兼容新旧格式
        scenario_obj = result.get('scenario', result)

        # 只处理多参数情景（名称不带"市场指数"、"行业指数"等前缀的）
        scenario_name = scenario_obj.get('name', '')
        if any(prefix in scenario_name for prefix in ['市场指数', '行业指数', '行业PE', '个股PE', 'DCF估值']):
            continue

        # 获取实际溢价率（百分比形式）
        premium_rate_pct = scenario_obj.get('premium_rate', scenario_obj.get('discount', 0)) * 100
        median_return_pct = result.get('median_return', 0) * 100
        # 修正：使用var_5（5%分位数）作为95% VaR，保留原始值（负值表示损失）
        var_5_raw = result.get('var_5', result.get('var_95', 0)) * 100  # 原始值，可能是负值
        var_95_pct = abs(var_5_raw)  # 绝对值，用于筛选条件
        profit_prob_pct = result.get('profit_prob', 0)

        # 检查四个条件
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
                'var_95': var_5_raw,  # 存储原始值（可能是负值），用于显示
                'profit_prob': profit_prob_pct
            })

    if qualified_585_scenarios:
        add_paragraph(document, f' 从585种情景中找到{len(qualified_585_scenarios)}个符合条件的方案', bold=True)
        add_paragraph(document, '')

        # 按收益率排序，选择最接近8%的方案
        qualified_585_scenarios.sort(key=lambda x: abs(x['median_return'] - 8))
        best_585_scenario = qualified_585_scenarios[0]

        add_paragraph(document, '推荐方案（最接近8%收益率）：', bold=True)
        add_paragraph(document, '')

        recommended_585_data = [
            ['建议报价', f"{best_585_scenario['issue_price']:.2f}元/股", f"溢价率{best_585_scenario['premium_rate']:+.1f}%"],
            ['预期中位数收益率', f"{best_585_scenario['median_return']:+.1f}%", '年化收益率'],
            ['盈利概率', f"{best_585_scenario['profit_prob']:.1f}%", '蒙特卡洛模拟盈利概率'],
            ['95% VaR', f"{best_585_scenario['var_95']:+.1f}%", '锁定期最大损失（负值）'],
            ['', '', ''],
            ['市场参数', f"漂移率{best_585_scenario['drift']*100:+.0f}%, 波动率{best_585_scenario['volatility']*100:.0f}%", '基于历史数据'],
            [' 溢价率', f"{best_585_scenario['premium_rate']:+.1f}%", '≤ -5% ✓'],
            [' 中位数收益率', f"{best_585_scenario['median_return']:+.1f}%", '> 8% ✓'],
            [' 95% VaR', f"{best_585_scenario['var_95']:+.1f}%", '|VaR| < 60% ✓'],
            [' 盈利概率', f"{best_585_scenario['profit_prob']:.1f}%", '> 50% ✓'],
        ]

        add_table_data(document, ['指标', '数值', '说明'], recommended_585_data)

        # 推荐理由
        add_paragraph(document, '')
        add_paragraph(document, '推荐理由：', bold=True)
        add_paragraph(document, f'• 该方案从585种情景中筛选而出，满足所有4个筛选条件')
        add_paragraph(document, f'• 中位数收益率{best_585_scenario["median_return"]:+.1f}%最接近8%目标，符合保守原则')
        add_paragraph(document, f'• 溢价率{best_585_scenario["premium_rate"]:+.1f}%提供{abs(best_585_scenario["premium_rate"]):.1f}%的安全边际')
        add_paragraph(document, f'• 盈利概率{best_585_scenario["profit_prob"]:.1f}%，超过半数场景盈利')
        # 修正：VaR是负值表示损失
        var_585_value = best_585_scenario["var_95"]
        var_585_desc = f"{abs(var_585_value):.1f}%的损失" if var_585_value < 0 else f"{var_585_value:.1f}%的收益"
        add_paragraph(document, f'• 95% VaR为{var_585_value:+.1f}%（最大可能{var_585_desc}），风险相对可控')

        # 投资建议
        add_paragraph(document, '')
        add_paragraph(document, ' 投资建议：', bold=True)

        if best_585_scenario['median_return'] >= 15:
            add_paragraph(document, f'•  强烈推荐 - 该方案预期收益率{best_585_scenario["median_return"]:+.1f}%较高，风险可控')
        elif best_585_scenario['median_return'] >= 10:
            add_paragraph(document, f'•  推荐 - 该方案预期收益率{best_585_scenario["median_return"]:+.1f}%达标，风险收益平衡')
        else:
            add_paragraph(document, f'•  谨慎推荐 - 该方案预期收益率{best_585_scenario["median_return"]:+.1f}%勉强达标，需密切关注风险')

        if best_585_scenario['profit_prob'] >= 70:
            add_paragraph(document, f'• 盈利概率{best_585_scenario["profit_prob"]:.1f}%较高，投资安全性较好')
        elif best_585_scenario['profit_prob'] >= 60:
            add_paragraph(document, f'• 盈利概率{best_585_scenario["profit_prob"]:.1f}%中等，需评估风险承受能力')
        else:
            add_paragraph(document, f'•  盈利概率{best_585_scenario["profit_prob"]:.1f}%偏低，建议提高安全边际要求')

        add_paragraph(document, '')

    else:
        add_paragraph(document, ' 从585种情景中未找到完全符合所有4个筛选条件的方案', bold=True)
        add_paragraph(document, '')
        add_paragraph(document, '建议：')
        add_paragraph(document, '• 适当降低预期收益率要求（如从8%降至5%）')
        add_paragraph(document, '• 要求更高的折价率（如-10%或更低）')
        add_paragraph(document, '• 调整报价策略，等待更好的市场时机')
        add_paragraph(document, '• 参考9.3.2.1的反向推算结果作为报价上限')

    add_paragraph(document, '')
    add_paragraph(document, '')

    # ==================== 9.4 当前发行价评估 ====================
    add_title(document, '9.4 当前发行价评估', level=2)
    add_paragraph(document, '本节评估当前发行价的合理性，并给出投资建议。')
    add_paragraph(document, '')

    # 计算预期价格（用于后续评估）
    historical_return = market_data.get('annual_return_120d', 0.08)
    expected_price = current_price_eval * (1 + historical_return * lockup_years)

    # 评估当前发行价
    current_premium = (issue_price_eval - current_price_eval) / current_price_eval * 100
    add_paragraph(document, f'• 当前发行价: {issue_price_eval:.2f} 元')
    add_paragraph(document, f'• 当前价格: {current_price_eval:.2f} 元')
    add_paragraph(document, f'• 溢价率（相对当前价）: {current_premium:+.2f}%')

    # 说明：溢价率为负表示折价，为正表示溢价
    if current_premium > 0:
        add_paragraph(document, f'  （溢价{current_premium:.2f}%，发行价高于当前价）')
    elif current_premium < 0:
        add_paragraph(document, f'  （折价{abs(current_premium):.2f}%，发行价低于当前价）')
    else:
        add_paragraph(document, f'  （平价发行，发行价等于当前价）')

    add_paragraph(document, f'• 溢价率（相对MA20: {ma20_mc:.2f}元）: {discount_premium:+.2f}%')
    add_paragraph(document, '')

    # 判断当前发行价是否合理
    max_price_20 = expected_price / (1 + 0.20 * lockup_years)

    if issue_price_eval <= max_price_20:
        add_paragraph(document, f'•  当前发行价低于20%目标收益对应的最高报价({max_price_20:.2f}元)，具有投资价值')
        investment_advice = "建议积极参与"
    elif issue_price_eval <= max_price_20 * 1.05:
        add_paragraph(document, f'•  当前发行价接近20%目标收益对应的最高报价({max_price_20:.2f}元)，需谨慎评估')
        investment_advice = "建议谨慎参与或要求更高折价"
    else:
        add_paragraph(document, f'•  当前发行价高于20%目标收益对应的最高报价({max_price_20:.2f}元)，安全边际不足')
        investment_advice = "建议规避或等待更好时机"

    add_paragraph(document, '')

    # ==================== 9.5 主要风险提示 ====================
    add_title(document, '9.5 主要风险提示', level=2)
    add_paragraph(document, '基于多维度风险分析，提示以下主要风险（详见前文各章节详细分析）：')
    add_paragraph(document, '')

    # 1. 市场风险
    add_paragraph(document, '1. 市场风险')
    add_paragraph(document, f'   • 波动率风险：当前120日窗口年化波动率为{mc_volatility_120d*100:.1f}%，市场波动可能导致实际收益偏离预期')
    add_paragraph(document, f'   • 趋势风险：当前120日窗口年化漂移率为{mc_drift_120d*100:+.2f}%，{"上升趋势" if mc_drift_120d > 0 else "下降趋势" if mc_drift_120d < 0 else "震荡趋势"}可能影响解禁时收益')
    add_paragraph(document, '')

    # 2. 流动性风险
    add_paragraph(document, '2. 流动性风险')
    add_paragraph(document, f'   • 锁定期风险：{project_params["lockup_period"]}个月锁定期内无法交易，需承担期间价格波动')
    add_paragraph(document, '   • 解禁冲击：解禁后可能面临抛压，导致实际变现价格低于理论价格')
    add_paragraph(document, '')

    # 3. VaR在险价值风险
    add_paragraph(document, '3. VaR在险价值风险')
    # 使用基于期收益率（锁定期收益率）的VaR，而不是年化VaR
    # 从context中获取var_results，然后获取120日窗口的var_95
    var_results = context.get('results', {}).get('var_results', {})
    var_95_period = var_results.get('120日', {}).get('var_95', 0.5)
    var_95_display = var_95_period * 100  # 转换为百分比
    add_paragraph(document, f'   • 120日窗口：95%置信水平下最大可能亏损{var_95_display:.1f}%')
    add_paragraph(document, '   • 尾部风险：历史数据显示，小概率极端事件（黑天鹅）可能导致损失超过VaR预测值')
    add_paragraph(document, '')

    # 4. 估值风险
    add_paragraph(document, '4. 估值风险')
    add_paragraph(document, f'   • DCF估值风险：DCF内在价值{intrinsic_value:.2f}元/股基于多个假设，实际业绩可能偏离预测')
    add_paragraph(document, '   • 相对估值风险：PE/PS/PB相对估值基于行业平均水平，行业景气度变化可能导致估值体系重构')
    add_paragraph(document, '')

    # 5. 发行定价风险
    add_paragraph(document, '5. 发行定价风险')
    if current_premium < 0:
        discount_amount = abs(current_premium)
        add_paragraph(document, f'   • 折价情况：当前溢价率为{current_premium:+.2f}%（折价{discount_amount:.2f}%）')
        if discount_amount < 10:
            add_paragraph(document, f'   • 折价不足风险：折价率低于10%，安全边际有限')
        else:
            add_paragraph(document, f'   • 折价合理：折价{discount_amount:.2f}%提供了一定的安全边际')
    else:
        add_paragraph(document, f'   • 溢价风险：当前溢价率为{current_premium:+.2f}%（溢价发行）')
        add_paragraph(document, f'   • 溢价发行无安全边际，需重点关注公司成长性')
    add_paragraph(document, '   • 定价偏离：若发行价显著高于盈亏平衡价格，将大幅降低盈利概率和预期收益')
    add_paragraph(document, '')

    # 6. 其他风险
    add_paragraph(document, '6. 其他风险')
    add_paragraph(document, '   • 行业政策风险：需关注行业监管政策变化')
    add_paragraph(document, '   • 业绩波动风险：需关注公司业绩预告、审计报告等')
    add_paragraph(document, '   • 竞争格局风险：行业竞争加剧可能影响盈利能力')
    add_paragraph(document, '')

    add_paragraph(document, '综合建议：')
    risk_advice_text = "风险较低，可积极参与" if total_score >= 80 else "风险中等，需谨慎评估" if total_score >= 60 else "风险较高，建议规避"
    add_paragraph(document, f'• 当前项目风险评分{total_score}/100分，{risk_advice_text}')
    add_paragraph(document, f'• 推荐方案：{investment_advice}')
    add_paragraph(document, '• 建议结合个人风险承受能力、资金成本和市场环境选择合适的报价方案')
    add_paragraph(document, '')
    if total_score >= 80 and issue_price_eval <= max_price_20:
        final_recommendation = f"🟢 建议积极参与"
        reason = f"风险评分{total_score}/100（优秀），发行价具有较好安全边际，符合20%目标收益率要求。"
    elif total_score >= 60:
        final_recommendation = f"🟡 谨慎参与"
        reason = f"风险评分{total_score}/100（中等），需结合溢价率和增长预期综合评估。"
    else:
        final_recommendation = f"🔴 建议规避"
        reason = f"风险评分{total_score}/100（较低），安全边际不足，风险较高。"

    add_paragraph(document, f'{final_recommendation}')
    add_paragraph(document, reason)

    add_paragraph(document, '')

    # ==================== 9.6 风控策略建议 ====================
    add_title(document, '9.6 风控策略建议', level=2)
    add_paragraph(document, '根据当前溢价率和盈利概率，提供风控策略建议。')
    add_paragraph(document, '')

    # 使用current_premium（溢价率），为负表示折价
    if current_premium <= -15:
        add_paragraph(document, '• 当前溢价率≤-15%（折价≥15%），可按原计划参与')
    elif current_premium <= -10:
        add_paragraph(document, '• 当前溢价率-10%~-15%（折价10%~15%），建议关注公司基本面和行业前景')
    else:
        add_paragraph(document, f'• 当前溢价率{current_premium:+.2f}%（折价不足或溢价），建议要求更高折价或等待更好时机')

    # 根据盈利概率给出建议（统一处理百分比格式）
    profit_prob_pct = profit_prob * 100 if profit_prob <= 1 else profit_prob
    if profit_prob_pct >= 70:
        add_paragraph(document, f'• 盈利概率{profit_prob_pct:.1f}%（≥70%），投资安全边际充足')
    elif profit_prob_pct >= 50:
        add_paragraph(document, f'• 盈利概率{profit_prob_pct:.1f}%（50%-70%），建议分批参与或控制仓位')
    else:
        add_paragraph(document, f'• 盈利概率{profit_prob_pct:.1f}%（<50%），风险较大，建议谨慎')

    add_paragraph(document, '')

    # ==================== 9.7 压力测试与VaR风险提示 ====================
    add_title(document, '9.7 压力测试与VaR风险提示', level=2)
    add_paragraph(document, '本节汇总压力测试和VaR分析的关键风险指标。')
    add_paragraph(document, '')

    # 从多窗口期VaR分析中提取关键风险指标
    var_95_pct = var_95 * 100
    var_99_pct = var_99 * 100 if 'var_99' in locals() else None

    add_paragraph(document, f'• 120日窗口VaR风险：95%置信水平下最大可能亏损为{var_95_pct:.1f}%')
    if var_99_pct:
        add_paragraph(document, f'• 极端情况VaR：99%置信水平下最大可能亏损为{var_99_pct:.1f}%')

    # 压力测试风险提示
    add_paragraph(document, '• 压力测试情景：需关注PE回归、市场危机、三重打击等极端情景下的潜在损失')
    add_paragraph(document, '• 波动率放大风险：当实际波动率超过120日窗口统计值时，风险敞口将显著增加')

    add_paragraph(document, '')

    # ==================== 9.8 多重方案选项建议 ====================
    add_title(document, '9.8 多重方案选项建议', level=2)
    add_paragraph(document, '根据不同的风险偏好，提供以下参与方案：')
    add_paragraph(document, '')

    add_paragraph(document, '根据不同的风险偏好，提供以下参与方案：')
    add_paragraph(document, '')

    # 方案一：保守型（低风险偏好）
    add_paragraph(document, '🛡️ 方案一：保守型（低风险偏好）')
    add_paragraph(document, f'   • 适用场景：追求确定性，可接受15%年化收益率')
    max_price_conservative = expected_price / (1 + 0.15 * lockup_years)
    discount_conservative = (current_price_eval - max_price_conservative) / current_price_eval * 100
    add_paragraph(document, f'   • 建议报价：不高于{max_price_conservative:.2f}元（折价率{discount_conservative:+.2f}%）')
    add_paragraph(document, f'   • 风险控制：要求较高折价，确保足够安全边际')
    add_paragraph(document, '')

    # 方案二：平衡型（中等风险偏好）
    add_paragraph(document, '⚖️ 方案二：平衡型（中等风险偏好）')
    add_paragraph(document, f'   • 适用场景：平衡收益与风险，目标20%年化收益率')
    max_price_balanced = expected_price / (1 + 0.20 * lockup_years)
    discount_balanced = (current_price_eval - max_price_balanced) / current_price_eval * 100
    add_paragraph(document, f'   • 建议报价：不高于{max_price_balanced:.2f}元（折价率{discount_balanced:+.2f}%）')
    add_paragraph(document, f'   • 风险控制：适度折价，关注盈利概率和VaR指标')
    add_paragraph(document, '')

    # 方案三：积极型（高风险偏好）
    add_paragraph(document, ' 方案三：积极型（高风险偏好）')
    add_paragraph(document, f'   • 适用场景：看好公司长期成长，可承受较大波动，目标25%年化收益率')
    max_price_aggressive = expected_price / (1 + 0.25 * lockup_years)
    discount_aggressive = (current_price_eval - max_price_aggressive) / current_price_eval * 100
    add_paragraph(document, f'   • 建议报价：不高于{max_price_aggressive:.2f}元（折价率{discount_aggressive:+.2f}%）')
    add_paragraph(document, f'   • 风险控制：需重点关注公司基本面和行业前景，建议控制仓位')
    add_paragraph(document, '')

    # 风险提示
    add_paragraph(document, ' 重要提示：')
    add_paragraph(document, '• 本建议基于120日窗口历史数据，实际收益可能偏离预期')
    add_paragraph(document, '• 市场环境变化、公司业绩波动等风险因素可能影响最终收益')
    add_paragraph(document, '• 建议结合最新市场情况和公司公告动态调整报价策略')
    add_paragraph(document, '• 投资有风险，决策需谨慎')

    add_paragraph(document, '')

    # ==================== 9.9 免责声明 ====================
    add_title(document, '9.9 免责声明', level=2)

    disclaimer = f'''
    本报告由自动化分析系统生成，仅供参考使用，不构成投资建议。

    1. 本报告基于历史数据和公开信息进行分析，不保证数据的准确性和完整性；
    2. 市场有风险，投资需谨慎。本报告中的任何分析观点不代表未来表现；
    3. 本报告提到的任何证券或投资标的，仅为分析示例，不构成推荐；
    4. 投资者应根据自身情况独立判断，自行承担投资风险；
    5. 本报告的知识产权归分析团队所有，未经许可不得转载或使用。

    报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    '''

    add_paragraph(document, disclaimer, font_size=9)

    add_section_break(document)

    return context
