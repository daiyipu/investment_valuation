# -*- coding: utf-8 -*-
"""
第六章 - 情景分析

功能：生成情景分析章节内容
"""

import sys
import os
import json
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# 添加路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)

from utils.font_manager import get_font_prop
from utils.analysis_tools import PrivatePlacementRiskAnalyzer

# 获取中文字体
font_prop = get_font_prop()


def add_title(document, text, level=1):
    """添加标题"""
    if level == 1:
        heading = document.add_heading(text, level=level)
        for run in heading.runs:
            run.font.name = '方正公文小标宋_GBK'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '方正公文小标宋_GBK')
            run.font.size = Pt(16)
    elif level == 2:
        heading = document.add_heading(text, level=level)
        for run in heading.runs:
            run.font.name = '方正公文小标宋_GBK'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '方正公文小标宋_GBK')
            run.font.size = Pt(15)
    else:
        heading = document.add_heading(text, level=level)
        for run in heading.runs:
            run.font.name = '黑体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
            run.font.size = Pt(14)
    return heading


def add_paragraph(document, text, bold=False, font_size=14):
    """添加段落"""
    para = document.add_paragraph(text)
    for run in para.runs:
        run.font.bold = bold
        run.font.size = Pt(font_size)
        if '<b>' in text and '</b>' in text:
            # 处理HTML格式的粗体标记
            run.text = text.replace('<b>', '').replace('</b>', '')
            run.font.bold = True
    return para


def add_table_data(document, headers, data, font_size=12):
    """添加表格"""
    table = document.add_table(rows=1, cols=len(headers))
    table.style = 'Light Grid Accent 1'

    # 设置表头
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        header_cells[i].text = header
        para = header_cells[i].paragraphs[0]
        para.runs[0].font.bold = True
        para.runs[0].font.size = Pt(font_size)
        para.runs[0].font.name = '宋体'
        para.runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 添加数据行
    for row_data in data:
        row_cells = table.add_row().cells
        for i, cell_data in enumerate(row_cells):
            row_cells[i].text = str(cell_data)
            para = row_cells[i].paragraphs[0]
            para.runs[0].font.size = Pt(font_size)
            para.runs[0].font.name = '宋体'
            para.runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            if isinstance(cell_data, (int, float)):
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    return table


def add_image(document, image_path, width=Inches(5)):
    """添加图片到文档"""
    if os.path.exists(image_path):
        document.add_picture(image_path, width=width)
        # 设置图片居中
        last_paragraph = document.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return True
    else:
        print(f"⚠️ 图片不存在: {image_path}")
        return False


def add_section_break(document):
    """添加分页符"""
    document.add_page_break()


def generate_multi_dimension_scenario_charts_split(current_price, base_price, volatility, drift, lockup_period, save_dir):
    """生成多维度情景图表 - 拆分版本"""
    from scipy import stats

    os.makedirs(save_dir, exist_ok=True)
    chart_paths = []

    # 定义参数范围
    drift_range = np.arange(-0.30, 0.35, 0.05)
    vol_range = np.arange(0.10, 0.55, 0.10)
    discount_range = np.arange(-0.20, 0.25, 0.05)

    all_results = []
    lockup_years = lockup_period / 12
    n_sim = 1000

    for d in drift_range:
        for v in vol_range:
            for disc in discount_range:
                issue_price = base_price * (1 + disc)
                lockup_drift = d * lockup_years
                lockup_vol = v * np.sqrt(lockup_years)

                np.random.seed(42)
                sim_returns = np.random.normal(lockup_drift, lockup_vol, n_sim)
                final_prices = current_price * np.exp(sim_returns)
                total_returns = (final_prices - issue_price) / issue_price
                annualized_returns = total_returns / lockup_years

                profit_prob = (total_returns > 0).mean() * 100
                mean_return = annualized_returns.mean() * 100

                all_results.append({
                    'drift': d,
                    'volatility': v,
                    'discount': disc,
                    'profit_prob': profit_prob,
                    'mean_return': mean_return
                })

    df = pd.DataFrame(all_results)

    # 图1: 不同波动率下的情景对比
    fig, ax = plt.subplots(figsize=(14, 7))
    vol_groups = df.groupby('volatility')

    for vol, group in vol_groups:
        ax.plot(group['drift'] * 100, group['profit_prob'],
               marker='o', label=f'波动率 {vol*100:.0f}%', linewidth=2, markersize=6)

    ax.set_xlabel('漂移率 (%)', fontproperties=font_prop, fontsize=13)
    ax.set_ylabel('盈利概率 (%)', fontproperties=font_prop, fontsize=13)
    ax.set_title('不同波动率下的盈利概率对比', fontproperties=font_prop, fontsize=15, fontweight='bold')
    ax.legend(prop=font_prop, fontsize=11)
    ax.grid(True, alpha=0.3)
    ax.axhline(y=50, color='red', linestyle='--', alpha=0.5, label='盈亏平衡线')

    for label in ax.get_xticklabels():
        label.set_fontproperties(font_prop)
        label.set_fontsize(11)
    for label in ax.get_yticklabels():
        label.set_fontproperties(font_prop)
        label.set_fontsize(11)

    plt.tight_layout()
    chart1_path = os.path.join(save_dir, '06_06_volatility_comparison.png')
    plt.savefig(chart1_path, dpi=150, bbox_inches='tight')
    plt.close()
    chart_paths.append(chart1_path)

    # 图2: 优质情景TOP10
    top_scenarios = df.nlargest(10, 'profit_prob')

    fig, ax = plt.subplots(figsize=(14, 7))
    colors = ['green' if p >= 70 else 'orange' if p >= 50 else 'red' for p in top_scenarios['profit_prob']]
    bars = ax.bar(range(len(top_scenarios)), top_scenarios['profit_prob'], color=colors, alpha=0.7, edgecolor='black')

    labels = [f"D:{row['drift']*100:+.0f}%, V:{row['volatility']*100:.0f}%, Dsc:{row['discount']*100:+.0f}%"
              for _, row in top_scenarios.iterrows()]
    ax.set_xticks(range(len(top_scenarios)))
    ax.set_xticklabels(labels, fontproperties=font_prop, fontsize=9, rotation=45, ha='right')
    ax.set_ylabel('盈利概率 (%)', fontproperties=font_prop, fontsize=13)
    ax.set_title('优质情景TOP10 (盈利概率排序)', fontproperties=font_prop, fontsize=15, fontweight='bold')
    ax.set_ylim(0, 100)
    ax.grid(True, alpha=0.3, axis='y')

    for i, (bar, prob) in enumerate(zip(bars, top_scenarios['profit_prob'])):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
               f'{prob:.1f}%', ha='center', va='bottom', fontproperties=font_prop, fontsize=10)

    for label in ax.get_yticklabels():
        label.set_fontproperties(font_prop)
        label.set_fontsize(11)

    plt.tight_layout()
    chart2_path = os.path.join(save_dir, '06_07_top_scenarios.png')
    plt.savefig(chart2_path, dpi=150, bbox_inches='tight')
    plt.close()
    chart_paths.append(chart2_path)

    return chart_paths


def generate_chapter(context):
    """
    生成第六章内容：情景分析

    参数:
        context: 包含所有必要数据的字典
            - document: Word文档对象
            - project_params: 项目参数
            - market_data: 市场数据
            - IMAGES_DIR: 图片保存目录
            - DATA_DIR: 数据目录
            - analyzer: 分析器对象
            - risk_params: 风险参数
            - stock_code: 股票代码
            - ma20: MA20价格
            - stock_pe_data: 个股PE数据（可选）
            - industry_pe_data: 行业PE数据（可选）
            - intrinsic_value: DCF内在价值（可选）
    """
    document = context['document']
    project_params = context['project_params']
    market_data = context['market_data']
    IMAGES_DIR = context['IMAGES_DIR']
    DATA_DIR = context.get('DATA_DIR', os.path.join(os.path.dirname(PROJECT_DIR), 'data'))
    analyzer = context['analyzer']
    risk_params = context.get('risk_params', {})
    stock_code = context.get('stock_code', '300735.SZ')
    ma20 = context.get('ma20', market_data.get('ma_20', project_params['current_price']))
    stock_pe_data = context.get('stock_pe_data', None)
    industry_pe_data = context.get('industry_pe_data', None)
    intrinsic_value = context.get('intrinsic_value', None)

    # ==================== 六、情景分析 ====================
    add_title(document, '六、情景分析', level=1)

    add_paragraph(document, '本章节分析定增项目在不同预设情景下的风险表现，包括多维度情景分析（预期期间收益率、净利润率、波动率、发行价折价）以及破发概率情景分析。')

    # ==================== 6.1 多参数情景分析 ====================
    add_paragraph(document, '')
    add_title(document, '6.1 多参数情景分析', level=2)

    add_paragraph(document, '本节通过穷举漂移率、波动率、溢价率三个参数的组合，全面分析不同市场环境下的定增项目收益预期。')
    add_paragraph(document, '')

    # ====== 参数空间定义 ======
    add_paragraph(document, '6.1.1 参数空间定义', bold=True, font_size=14)
    add_paragraph(document, '通过以下参数的穷举组合，模拟585种不同市场情景：')

    # 定义参数范围
    drift_rates = np.arange(-0.30, 0.35, 0.05)  # -30% 到 +30%，每档5%
    volatilities = np.arange(0.10, 0.55, 0.10)  # 10% 到 50%，每档10%
    discounts = np.arange(-0.20, 0.25, 0.05)  # -20% 到 +20%，每档5%

    # 参数说明表格
    param_data = [
        ['参数', '范围', '步长', '组合数'],
        ['漂移率（年化收益率）', '-30% ~ +30%', '5%', len(drift_rates)],
        ['波动率（年化）', '10% ~ 50%', '10%', len(volatilities)],
        ['折价率', '-20% ~ +20%', '5%', len(discounts)],
        ['总组合数', '-', '-', f'{len(drift_rates) * len(volatilities) * len(discounts)}']
    ]
    add_table_data(document, ['参数', '范围', '步长', '组合数'], param_data)

    add_paragraph(document, '')
    add_paragraph(document, '参数说明：')
    add_paragraph(document, f'• 当前价格: {project_params["current_price"]:.2f} 元/股')
    add_paragraph(document, f'• MA20价格: {market_data.get("ma_20", 0):.2f} 元/股（作为发行定价基准）')
    add_paragraph(document, f'• 锁定期: {project_params["lockup_period"]} 个月')
    add_paragraph(document, f'• 漂移率: 反映股价的预期趋势（负值=下跌，正值=上涨）')
    add_paragraph(document, f'• 波动率: 反映股价的不确定性（越高=风险越大）')
    add_paragraph(document, f'• 溢价率: 发行价相对MA20的溢价（负值=折价，正值=溢价，与配置的premium_rate一致）')
    add_paragraph(document, '')

    # ====== 模拟所有组合 ======
    print("\n运行多参数组合模拟...")

    all_scenarios = []
    n_sim = 2000  # 每个组合模拟2000次
    lockup_years = project_params['lockup_period'] / 12
    current_price = project_params['current_price']
    ma20_price = market_data.get('ma_20', current_price)  # 使用MA20作为发行定价基准

    scenario_count = 0
    for drift in drift_rates:
        for vol in volatilities:
            for discount in discounts:
                scenario_count += 1

                # 计算发行价（使用MA20作为基准，与配置保持一致）
                issue_price = ma20_price * (1 + discount)

                # 蒙特卡洛模拟
                lockup_drift = drift * lockup_years
                lockup_vol = vol * np.sqrt(lockup_years)

                np.random.seed(42)  # 固定种子以确保可重复性
                sim_returns = np.random.normal(lockup_drift, lockup_vol, n_sim)
                final_prices = current_price * np.exp(sim_returns)

                # 计算收益
                total_returns = (final_prices - issue_price) / issue_price
                # 年化收益率（使用单利计算）
                annualized_returns = total_returns / lockup_years

                # 统计指标
                mean_return = annualized_returns.mean()
                median_return = np.median(annualized_returns)
                profit_prob = (total_returns > 0).mean() * 100
                var_5 = np.percentile(annualized_returns, 5)
                var_95 = np.percentile(annualized_returns, 95)

                all_scenarios.append({
                    'drift': drift,
                    'volatility': vol,
                    'discount': discount,
                    'issue_price': issue_price,
                    'mean_return': mean_return,
                    'median_return': median_return,
                    'profit_prob': profit_prob,
                    'var_5': var_5,
                    'var_95': var_95
                })

                if scenario_count % 100 == 0:
                    print(f"  已完成 {scenario_count}/{len(drift_rates) * len(volatilities) * len(discounts)} 个情景...")

    print(f"  完成！共模拟 {len(all_scenarios)} 个情景")

    # ====== 按收益和概率排序 ======
    add_paragraph(document, '')
    add_paragraph(document, '6.1.2 情景分析结果', bold=True, font_size=14)

    # 按预期年化收益排序（Top 20）
    add_paragraph(document, '表1: 按预期年化收益率排序的Top 20情景')
    add_paragraph(document, '')

    top_by_return = sorted(all_scenarios, key=lambda x: x['mean_return'], reverse=True)[:20]

    top_return_data = []
    for i, s in enumerate(top_by_return, 1):
        top_return_data.append([
            f"{i}",
            f"{s['drift']*100:+.0f}%",
            f"{s['volatility']*100:.0f}%",
            f"{s['discount']*100:+.0f}%",
            f"{s['issue_price']:.2f} 元",
            f"{s['mean_return']*100:+.2f}%",
            f"{s['profit_prob']:.1f}%"
        ])

    return_headers = ['排名', '漂移率', '波动率', '溢价率', '发行价', '预期年化收益', '盈利概率']
    add_table_data(document, return_headers, top_return_data)

    # 按盈利概率排序（Top 20）
    add_paragraph(document, '')
    add_paragraph(document, '表2: 按盈利概率排序的Top 20情景')
    add_paragraph(document, '')

    top_by_prob = sorted(all_scenarios, key=lambda x: x['profit_prob'], reverse=True)[:20]

    top_prob_data = []
    for i, s in enumerate(top_by_prob, 1):
        top_prob_data.append([
            f"{i}",
            f"{s['drift']*100:+.0f}%",
            f"{s['volatility']*100:.0f}%",
            f"{s['discount']*100:+.0f}%",
            f"{s['issue_price']:.2f} 元",
            f"{s['mean_return']*100:+.2f}%",
            f"{s['profit_prob']:.1f}%"
        ])

    add_table_data(document, return_headers, top_prob_data)

    # ====== 综合最优情景 ======
    add_paragraph(document, '')
    add_paragraph(document, '6.1.3 综合最优情景分析', bold=True, font_size=14)

    # 找出同时满足高收益和高概率的情景
    high_return = [s for s in all_scenarios if s['mean_return'] > 0.20]  # 年化收益>20%
    high_prob = [s for s in all_scenarios if s['profit_prob'] > 80]  # 盈利概率>80%
    best_scenarios = [s for s in high_return if s in high_prob]

    if best_scenarios:
        add_paragraph(document, f'找到 {len(best_scenarios)} 个同时满足"年化收益>20%"和"盈利概率>80%"的最优情景：')
        add_paragraph(document, '')

        best_data = []
        for i, s in enumerate(sorted(best_scenarios, key=lambda x: x['mean_return'], reverse=True)[:10], 1):
            best_data.append([
                f"{i}",
                f"{s['drift']*100:+.0f}%",
                f"{s['volatility']*100:.0f}%",
                f"{s['discount']*100:+.0f}%",
                f"{s['mean_return']*100:+.2f}%",
                f"{s['profit_prob']:.1f}%"
            ])
        add_table_data(document, ['排名', '漂移率', '波动率', '溢价率', '预期年化收益', '盈利概率'], best_data)
    else:
        add_paragraph(document, '未找到同时满足"年化收益>20%"和"盈利概率>80%"的情景。')

    # ====== 当前项目定位 ======
    add_paragraph(document, '')
    add_paragraph(document, '6.1.4 当前项目在所有情景中的定位', bold=True, font_size=14)

    current_drift = market_data.get('annual_return_120d', risk_params.get('drift', 0.08))
    current_vol = market_data.get('volatility_120d', risk_params.get('volatility', 0.30))
    # 使用MA20作为基准计算折价率（与配置保持一致）
    ma20_for_discount = market_data.get('ma_20', project_params['current_price'])
    current_discount = (project_params['issue_price'] / ma20_for_discount) - 1

    # 找到最接近当前参数的情景
    for s in all_scenarios:
        s['distance'] = (
            abs(s['drift'] - current_drift) +
            abs(s['volatility'] - current_vol) +
            abs(s['discount'] - current_discount)
        )

    closest_scenario = min(all_scenarios, key=lambda x: x['distance'])

    # 计算当前项目的排名
    return_rank = sorted(all_scenarios, key=lambda x: x['mean_return'], reverse=True).index(closest_scenario) + 1
    prob_rank = sorted(all_scenarios, key=lambda x: x['profit_prob'], reverse=True).index(closest_scenario) + 1

    add_paragraph(document, '基于当前市场参数的最接近情景：')
    add_paragraph(document, f'• 漂移率: {current_drift*100:+.2f}%')
    add_paragraph(document, f'• 波动率: {current_vol*100:.2f}%')
    add_paragraph(document, f'• 溢价率: {current_discount*100:+.2f}% （相对MA20: {ma20_for_discount:.2f}元）')
    add_paragraph(document, f'• 发行价: {project_params["issue_price"]:.2f} 元/股')
    add_paragraph(document, f'• 预期年化收益（算术平均）: {closest_scenario["mean_return"]*100:+.2f}%')
    add_paragraph(document, f'• 收益率中位数: {closest_scenario["median_return"]*100:+.2f}%')
    add_paragraph(document, f'• 盈利概率: {closest_scenario["profit_prob"]:.1f}%')
    add_paragraph(document, f'• 收益率排名: 第 {return_rank} 名 / 共 {len(all_scenarios)} 个情景')
    add_paragraph(document, f'• 盈利概率排名: 第 {prob_rank} 名 / 共 {len(all_scenarios)} 个情景')

    # =====% 分析结论 ======
    add_paragraph(document, '')
    add_paragraph(document, '6.1.5 分析结论', bold=True, font_size=14)

    # 统计分析
    avg_return_all = np.mean([s['mean_return'] for s in all_scenarios])
    avg_prob_all = np.mean([s['profit_prob'] for s in all_scenarios])

    profit_scenarios = [s for s in all_scenarios if s['profit_prob'] >= 50]
    loss_scenarios = [s for s in all_scenarios if s['profit_prob'] < 50]

    add_paragraph(document, f'• 在全部{len(all_scenarios)}个情景组合中：')
    add_paragraph(document, f"  - 盈利概率≥50%的情景: {len(profit_scenarios)} 个 ({len(profit_scenarios)/len(all_scenarios)*100:.1f}%)")
    add_paragraph(document, f"  - 盈利概率<50%的情景: {len(loss_scenarios)} 个 ({len(loss_scenarios)/len(all_scenarios)*100:.1f}%)")

    add_paragraph(document, '')
    add_paragraph(document, f'• 平均预期年化收益: {avg_return_all*100:+.2f}%')
    add_paragraph(document, f'• 平均盈利概率: {avg_prob_all:.1f}%')

    add_paragraph(document, '')
    add_paragraph(document, '关键发现：')

    # 找出漂移率的影响
    pos_drift_scenarios = [s for s in all_scenarios if s['drift'] >= 0]
    neg_drift_scenarios = [s for s in all_scenarios if s['drift'] < 0]

    if pos_drift_scenarios and neg_drift_scenarios:
        avg_return_pos = np.mean([s['mean_return'] for s in pos_drift_scenarios])
        avg_return_neg = np.mean([s['mean_return'] for s in neg_drift_scenarios])
        avg_prob_pos = np.mean([s['profit_prob'] for s in pos_drift_scenarios])
        avg_prob_neg = np.mean([s['profit_prob'] for s in neg_drift_scenarios])

        add_paragraph(document, f'• 漂移率对收益影响显著：')
        add_paragraph(document, f"  - 正漂移率情景平均收益: {avg_return_pos*100:+.2f}%, 盈利概率: {avg_prob_pos:.1f}%")
        add_paragraph(document, f"  - 负漂移率情景平均收益: {avg_return_neg*100:+.2f}%, 盈利概率: {avg_prob_neg:.1f}%")

    # 找出折价率的影响
    deep_discount_scenarios = [s for s in all_scenarios if s['discount'] <= -0.15]
    premium_scenarios = [s for s in all_scenarios if s['discount'] > 0]

    if deep_discount_scenarios and premium_scenarios:
        avg_return_discount = np.mean([s['mean_return'] for s in deep_discount_scenarios])
        avg_return_premium = np.mean([s['mean_return'] for s in premium_scenarios])
        avg_prob_discount = np.mean([s['profit_prob'] for s in deep_discount_scenarios])
        avg_prob_premium = np.mean([s['profit_prob'] for s in premium_scenarios])

        add_paragraph(document, '')
        add_paragraph(document, f'• 溢价率是盈利概率的关键：')
        add_paragraph(document, f"  - 深度折价(≤-15%)情景平均收益: {avg_return_discount*100:+.2f}%, 盈利概率: {avg_prob_discount:.1f}%")
        add_paragraph(document, f"  - 溢价情景平均收益: {avg_return_premium*100:+.2f}%, 盈利概率: {avg_prob_premium:.1f}%")

    add_paragraph(document, '')
    add_paragraph(document, '投资建议：')
    if current_drift < 0:
        add_paragraph(document, f'⚠️ 当前漂移率为{current_drift*100:+.2f}%（负值），建议要求较高折价（更负的溢价率）以补偿下行风险')
    if current_discount > -0.10:
        add_paragraph(document, f'⚠️ 当前溢价率仅为{current_discount*100:+.2f}%，建议提高至-15%以下（更深的折价）')
    else:
        add_paragraph(document, f'✅ 当前溢价率{current_discount*100:+.2f}%较为合理，提供了一定的安全边际')

    # 保存all_scenarios供附件使用
    all_scenarios_for_appendix = all_scenarios.copy()

    # ==================== 6.2 基于市场指数与行业的情景分析 ====================
    add_paragraph(document, '')
    add_title(document, '6.2 基于市场指数与行业的情景分析', level=2)

    add_paragraph(document, '本节基于市场指数和行业的历史数据，构建多个典型情景进行蒙特卡洛模拟分析，评估不同市场环境下的定增收益。')
    add_paragraph(document, '通过分析当前情景、乐观情景、中性情景和悲观情景，全面评估项目在各种市场条件下的风险收益特征。')
    add_paragraph(document, '')
    add_paragraph(document, '本节基于行业和指数的历史数据，构建6个典型情景进行蒙特卡洛模拟分析，评估不同市场环境下的定增收益。')
    add_paragraph(document, '')

    # 重新加载行业数据（确保可用）
    industry_data_available = False
    industry_vol_120d = 0.35
    industry_return_120d = 0.08

    try:
        industry_data_file = os.path.join(DATA_DIR, f'{stock_code.replace(".", "_")}_industry_data.json')
        if os.path.exists(industry_data_file):
            with open(industry_data_file, 'r', encoding='utf-8') as f:
                industry_data_loaded = json.load(f)

            industry_vol_120d = industry_data_loaded.get('volatility_120d', 0.35)
            industry_return_120d = industry_data_loaded.get('annual_return_120d', 0.08)
            industry_data_available = True
            print(f"✅ 已重新加载行业数据: {industry_data_file}")
    except Exception as e:
        print(f"⚠️ 无法加载行业数据，使用默认值: {e}")

    # 加载指数数据
    indices_data_available = False
    index_vol_120d = 0.30
    index_return_120d = 0.08

    try:
        indices_data_file = os.path.join(DATA_DIR, 'market_indices_scenario_data_v2.json')
        if os.path.exists(indices_data_file):
            with open(indices_data_file, 'r', encoding='utf-8') as f:
                indices_data_loaded = json.load(f)

            # 计算主要指数的平均值（沪深300、中证500、创业板指、科创50）
            major_indices = ['沪深300', '中证500', '创业板指', '科创50']
            vol_values = []
            return_values = []

            for index_name in major_indices:
                if index_name in indices_data_loaded:
                    vol_values.append(indices_data_loaded[index_name]['volatility_120d'])
                    return_values.append(indices_data_loaded[index_name]['return_120d'])

            if vol_values and return_values:
                index_vol_120d = np.mean(vol_values)
                index_return_120d = np.mean(return_values)
                indices_data_available = True
                print(f"✅ 已加载指数数据: {indices_data_file}")
    except Exception as e:
        print(f"⚠️ 无法加载指数数据，使用默认值: {e}")

    # 计算当前项目的溢价率
    current_premium_scenario = (project_params['issue_price'] - project_params['current_price']) / project_params['current_price']

    # 定义情景参数（基于指数和行业的综合数据）
    scenarios_config = []

    # 情景1: 当前情景（标的股票120日窗口真实数据）
    scenarios_config.append({
        'name': '当前情景',
        'description': '标的股票120日窗口真实数据',
        'volatility': market_data.get('volatility_120d', 0.35),
        'drift': market_data.get('annual_return_120d', 0.08),
        'premium_rate': current_premium_scenario  # 使用刚才计算的溢价率
    })

    # 计算指数和行业的综合统计值
    # 情景2: 乐观情景（指数和行业中较高的漂移率、较低的波动率）
    if indices_data_available and industry_data_available:
        # 从指数和行业中取较高的漂移率
        optimistic_drift = max(index_return_120d, industry_return_120d)
        # 从指数和行业中取较低的波动率
        optimistic_vol = min(index_vol_120d, industry_vol_120d)

        scenarios_config.append({
            'name': '情景1（乐观）',
            'description': f'指数和行业120日窗口中漂移率较高值（{optimistic_drift*100:+.2f}%）、波动率较低值（{optimistic_vol*100:.2f}%）、溢价率-20%（深折价）',
            'volatility': optimistic_vol * 0.9,  # 取较低值后再打9折
            'drift': optimistic_drift * 1.1,  # 取较高值后再打1.1倍
            'premium_rate': -0.20  # -20%溢价率（深折价）
        })

        # 情景3: 中性情景（指数和行业中档值）
        neutral_drift = (index_return_120d + industry_return_120d) / 2
        neutral_vol = (index_vol_120d + industry_vol_120d) / 2

        scenarios_config.append({
            'name': '情景2（中性）',
            'description': f'指数和行业120日窗口中漂移率和波动率中档值（漂移率{neutral_drift*100:+.2f}%、波动率{neutral_vol*100:.2f}%）、溢价率-10%（折价）',
            'volatility': neutral_vol,
            'drift': neutral_drift,
            'premium_rate': -0.10  # -10%溢价率（折价）
        })

        # 情景4: 悲观情景（指数和行业中较低的漂移率、较高的波动率）
        pessimistic_drift = min(index_return_120d, industry_return_120d)
        pessimistic_vol = max(index_vol_120d, industry_vol_120d)

        scenarios_config.append({
            'name': '情景3（悲观）',
            'description': f'指数和行业120日窗口中漂移率较低值（{pessimistic_drift*100:+.2f}%）、波动率较高值（{pessimistic_vol*100:.2f}%）、溢价率0%（平价）',
            'volatility': pessimistic_vol * 1.1,  # 取较高值后再打1.1倍
            'drift': pessimistic_drift * 0.9,  # 取较低值后再打9折
            'premium_rate': 0.00  # 0%溢价率（平价）
        })
    else:
        # 如果没有指数数据，跳过指数情景分析，继续执行后续PE分位数分析
        print(f"⚠️ 指数数据不可用，跳过指数情景分析")

    # 情景4-6: 基于行业PE分位数的情景（如果有PE数据）
    pe_scenarios = []
    if stock_pe_data is not None and industry_pe_data is not None:
        try:
            # 计算行业PE分位数
            industry_pe_values = industry_pe_data['pe_ttm'].dropna()
            pe_25 = industry_pe_values.quantile(0.25)
            pe_50 = industry_pe_values.quantile(0.50)
            pe_75 = industry_pe_values.quantile(0.75)

            # 获取标的个股当前PE值
            stock_pe_current = stock_pe_data['pe_ttm'].iloc[-1]

            # 根据行业PE分位数与标的个股PE的比值计算漂移率
            # 漂移率 = (行业PE分位数 / 个股PE - 1)
            # 这表示：如果个股PE相对行业PE更低，估值便宜，未来增长率应该更高

            # 计算各分位数的比值
            ratio_75 = pe_75 / stock_pe_current if stock_pe_current > 0 else 1.0
            ratio_50 = pe_50 / stock_pe_current if stock_pe_current > 0 else 1.0
            ratio_25 = pe_25 / stock_pe_current if stock_pe_current > 0 else 1.0

            # 漂移率 = 比值 - 1（直接作为增长率假设）
            drift_pe_75 = ratio_75 - 1.0
            drift_pe_50 = ratio_50 - 1.0
            drift_pe_25 = ratio_25 - 1.0

            # 限制漂移率在合理范围内（-50%到+100%）
            drift_pe_75 = max(-0.50, min(1.00, drift_pe_75))
            drift_pe_50 = max(-0.50, min(1.00, drift_pe_50))
            drift_pe_25 = max(-0.50, min(1.00, drift_pe_25))

            print(f"✅ PE分位数漂移率计算完成:")
            print(f"   标的个股当前PE: {stock_pe_current:.2f}倍")
            print(f"   行业PE 75%分位数: {pe_75:.1f}倍, 比值={ratio_75:.2f}, 漂移率={drift_pe_75*100:+.2f}%")
            print(f"   行业PE 50%分位数: {pe_50:.1f}倍, 比值={ratio_50:.2f}, 漂移率={drift_pe_50*100:+.2f}%")
            print(f"   行业PE 25%分位数: {pe_25:.1f}倍, 比值={ratio_25:.2f}, 漂移率={drift_pe_25*100:+.2f}%")
            print(f"   计算公式: 漂移率 = (行业PE分位数 / 个股PE - 1)")

            # 情景4: 行业PE 75%分位数 + 深折价
            scenarios_config.append({
                'name': '情景4（行业PE 75%分位数）',
                'description': f'行业PE 75%分位数({pe_75:.1f}倍)与个股PE({stock_pe_current:.2f}倍)比值{ratio_75:.2f}，漂移率{drift_pe_75*100:+.2f}%，溢价率-20%（深折价）',
                'volatility': industry_vol_120d,
                'drift': drift_pe_75,
                'premium_rate': -0.20,
                'pe_based': True
            })

            # 情景5: 行业PE 50%分位数 + 中等折价
            scenarios_config.append({
                'name': '情景5（行业PE 50%分位数）',
                'description': f'行业PE 50%分位数({pe_50:.1f}倍)与个股PE({stock_pe_current:.2f}倍)比值{ratio_50:.2f}，漂移率{drift_pe_50*100:+.2f}%，溢价率-10%（折价）',
                'volatility': industry_vol_120d,
                'drift': drift_pe_50,
                'premium_rate': -0.10,
                'pe_based': True
            })

            # 情景6: 行业PE 25%分位数 + 平价
            scenarios_config.append({
                'name': '情景6（行业PE 25%分位数）',
                'description': f'行业PE 25%分位数({pe_25:.1f}倍)与个股PE({stock_pe_current:.2f}倍)比值{ratio_25:.2f}，漂移率{drift_pe_25*100:+.2f}%，溢价率0%（平价）',
                'volatility': industry_vol_120d,
                'drift': drift_pe_25,
                'premium_rate': 0.00,
                'pe_based': True
            })

            pe_scenarios = ['情景4（行业PE 75%分位数）', '情景5（行业PE 50%分位数）', '情景6（行业PE 25%分位数）']

            # 情景7-9: 基于个股PE分位数的情景
            # 计算个股PE分位数
            stock_pe_values = stock_pe_data['pe_ttm'].dropna()
            stock_pe_25 = stock_pe_values.quantile(0.25)
            stock_pe_50 = stock_pe_values.quantile(0.50)
            stock_pe_75 = stock_pe_values.quantile(0.75)

            # 根据个股PE分位数与个股当前PE的比值计算漂移率
            # 漂移率 = (个股PE分位数 / 个股当前PE - 1)
            stock_ratio_75 = stock_pe_75 / stock_pe_current if stock_pe_current > 0 else 1.0
            stock_ratio_50 = stock_pe_50 / stock_pe_current if stock_pe_current > 0 else 1.0
            stock_ratio_25 = stock_pe_25 / stock_pe_current if stock_pe_current > 0 else 1.0

            stock_drift_75 = stock_ratio_75 - 1.0
            stock_drift_50 = stock_ratio_50 - 1.0
            stock_drift_25 = stock_ratio_25 - 1.0

            # 限制漂移率在合理范围内
            stock_drift_75 = max(-0.50, min(1.00, stock_drift_75))
            stock_drift_50 = max(-0.50, min(1.00, stock_drift_50))
            stock_drift_25 = max(-0.50, min(1.00, stock_drift_25))

            print(f"✅ 个股PE分位数漂移率计算完成:")
            print(f"   个股当前PE: {stock_pe_current:.2f}倍")
            print(f"   个股PE 75%分位数: {stock_pe_75:.1f}倍, 比值={stock_ratio_75:.2f}, 漂移率={stock_drift_75*100:+.2f}%")
            print(f"   个股PE 50%分位数: {stock_pe_50:.1f}倍, 比值={stock_ratio_50:.2f}, 漂移率={stock_drift_50*100:+.2f}%")
            print(f"   个股PE 25%分位数: {stock_pe_25:.1f}倍, 比值={stock_ratio_25:.2f}, 漂移率={stock_drift_25*100:+.2f}%")
            print(f"   计算公式: 漂移率 = (个股PE分位数 / 个股当前PE - 1)")

            # 情景7: 个股PE 75%分位数 + 深折价
            scenarios_config.append({
                'name': '情景7（个股PE 75%分位数）',
                'description': f'个股PE 75%分位数({stock_pe_75:.1f}倍)与当前PE({stock_pe_current:.2f}倍)比值{stock_ratio_75:.2f}，漂移率{stock_drift_75*100:+.2f}%，溢价率-20%（深折价）',
                'volatility': industry_vol_120d,
                'drift': stock_drift_75,
                'premium_rate': -0.20,
                'stock_pe_based': True
            })

            # 情景8: 个股PE 50%分位数 + 中等折价
            scenarios_config.append({
                'name': '情景8（个股PE 50%分位数）',
                'description': f'个股PE 50%分位数({stock_pe_50:.1f}倍)与当前PE({stock_pe_current:.2f}倍)比值{stock_ratio_50:.2f}，漂移率{stock_drift_50*100:+.2f}%，溢价率-10%（折价）',
                'volatility': industry_vol_120d,
                'drift': stock_drift_50,
                'premium_rate': -0.10,
                'stock_pe_based': True
            })

            # 情景9: 个股PE 25%分位数 + 平价
            scenarios_config.append({
                'name': '情景9（个股PE 25%分位数）',
                'description': f'个股PE 25%分位数({stock_pe_25:.1f}倍)与当前PE({stock_pe_current:.2f}倍)比值{stock_ratio_25:.2f}，漂移率{stock_drift_25*100:+.2f}%，溢价率0%（平价）',
                'volatility': industry_vol_120d,
                'drift': stock_drift_25,
                'premium_rate': 0.00,
                'stock_pe_based': True
            })

            pe_scenarios.extend(['情景7（个股PE 75%分位数）', '情景8（个股PE 50%分位数）', '情景9（个股PE 25%分位数）'])

            # 情景10-12: 基于DCF估值的情景
            # 检查是否有intrinsic_value（DCF内在价值）
            if intrinsic_value is not None:
                current_price_dcf = project_params['current_price']

                # 计算DCF估值与当前价格的比值
                dcf_ratio = intrinsic_value / current_price_dcf if current_price_dcf > 0 else 1.0

                # DCF估值的漂移率 = (DCF内在价值 / 当前价格 - 1)
                # 这表示如果DCF内在价值高于当前价格，未来收益率应该为正
                dcf_drift = dcf_ratio - 1.0

                # 限制漂移率在合理范围内
                dcf_drift = max(-0.50, min(1.00, dcf_drift))

                print(f"✅ DCF估值漂移率计算完成:")
                print(f"   DCF内在价值: {intrinsic_value:.2f}元/股")
                print(f"   当前价格: {current_price_dcf:.2f}元/股")
                print(f"   比值: {dcf_ratio:.2f}")
                print(f"   漂移率: {dcf_drift*100:+.2f}%")
                print(f"   计算公式: 漂移率 = (DCF内在价值 / 当前价格 - 1)")

                # 情景10: DCF估值 + 深折价
                scenarios_config.append({
                    'name': '情景10（DCF估值）',
                    'description': f'DCF内在价值({intrinsic_value:.2f}元)与当前价格({current_price_dcf:.2f}元)比值{dcf_ratio:.2f}，漂移率{dcf_drift*100:+.2f}%，溢价率-20%（深折价）',
                    'volatility': industry_vol_120d,
                    'drift': dcf_drift,
                    'premium_rate': -0.20,
                    'dcf_based': True
                })

                # 情景11: DCF估值 + 中等折价
                scenarios_config.append({
                    'name': '情景11（DCF估值）',
                    'description': f'DCF内在价值({intrinsic_value:.2f}元)与当前价格({current_price_dcf:.2f}元)比值{dcf_ratio:.2f}，漂移率{dcf_drift*100:+.2f}%，溢价率-10%（折价）',
                    'volatility': industry_vol_120d,
                    'drift': dcf_drift,
                    'premium_rate': -0.10,
                    'dcf_based': True
                })

                # 情景12: DCF估值 + 平价
                scenarios_config.append({
                    'name': '情景12（DCF估值）',
                    'description': f'DCF内在价值({intrinsic_value:.2f}元)与当前价格({current_price_dcf:.2f}元)比值{dcf_ratio:.2f}，漂移率{dcf_drift*100:+.2f}%，溢价率0%（平价）',
                    'volatility': industry_vol_120d,
                    'drift': dcf_drift,
                    'premium_rate': 0.00,
                    'dcf_based': True
                })

                pe_scenarios.extend(['情景10（DCF估值）', '情景11（DCF估值）', '情景12（DCF估值）'])
            else:
                print(f"⚠️ 未找到DCF内在价值数据，跳过DCF相关情景")
        except Exception as e:
            print(f"⚠️ PE数据计算失败，跳过PE相关情景: {e}")

    # 运行蒙特卡洛模拟计算每个情景的指标
    lockup_years = project_params['lockup_period'] / 12
    comprehensive_results = []

    for i, scenario in enumerate(scenarios_config):
        try:
            # 计算该情景的发行价
            scenario_issue_price = project_params['current_price'] * (1 + scenario['premium_rate'])

            # 运行蒙特卡洛模拟
            sim_scenario = analyzer.monte_carlo_simulation(
                n_simulations=2000,
                time_steps=120,  # 使用120日窗口
                volatility=scenario['volatility'],
                drift=scenario['drift'],
                seed=42
            )

            final_prices = sim_scenario.iloc[:, -1].values
            returns = (final_prices - scenario_issue_price) / scenario_issue_price
            annualized_returns = returns * (12 / project_params['lockup_period'])

            # 计算指标
            profit_prob = (returns > 0).mean() * 100
            median_return = np.median(annualized_returns) * 100
            mean_return = np.mean(annualized_returns) * 100
            var_95 = np.percentile(annualized_returns, 5) * 100

            comprehensive_results.append({
                'scenario': scenario,
                'issue_price': scenario_issue_price,
                'profit_prob': profit_prob,
                'median_return': median_return,
                'mean_return': mean_return,
                'var_95': var_95
            })

            # 调试输出
            print(f"✅ {scenario['name']}:")
            print(f"   发行价: {scenario_issue_price:.2f}元 (溢价率{scenario['premium_rate']*100:+.1f}%)")
            print(f"   漂移率: {scenario['drift']*100:+.2f}% (年化), 波动率: {scenario['volatility']*100:.2f}%")
            print(f"   盈利概率: {profit_prob:.1f}%, 中位数收益: {median_return:+.1f}%")

        except Exception as e:
            print(f"⚠️ {scenario['name']} 模拟失败: {e}")
            continue

    # 存储comprehensive_results到context，供第九章使用
    context['results']['comprehensive_results'] = comprehensive_results
    print(f"✅ 已保存comprehensive_results到context，共{len(comprehensive_results)}个情景")

    # 筛选6.2节相关情景（当前情景、情景1-3）
    index_industry_results = [r for r in comprehensive_results
                              if 'scenario' in r and
                              (r['scenario']['name'] == '当前情景' or
                               r['scenario']['name'].startswith('情景1（') or
                               r['scenario']['name'].startswith('情景2（') or
                               r['scenario']['name'].startswith('情景3（'))]

    # 生成6.2节的情景参数表格和图表
    if index_industry_results:
        add_paragraph(document, '')
        add_paragraph(document, '市场指数与行业情景参数表：')
        index_table_data = []
        for result in index_industry_results:
            scenario_obj = result['scenario']
            index_table_data.append([
                scenario_obj['name'],
                scenario_obj['description'],
                f"{scenario_obj['volatility']*100:.2f}%",
                f"{scenario_obj['drift']*100:+.2f}%",
                f"{scenario_obj['premium_rate']*100:+.0f}%",
                f"{result['profit_prob']:.1f}%",
                f"{result['median_return']:+.1f}%"
            ])
        add_table_data(document, ['情景名称', '情景描述', '波动率', '漂移率', '溢价率', '盈利概率', '收益率中位数'], index_table_data, font_size=10.5)

        # 生成情景对比图表
        try:
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

            scenario_names = [r['scenario']['name'] for r in index_industry_results]
            profit_probs = [r['profit_prob'] for r in index_industry_results]
            median_returns = [r['median_return'] for r in index_industry_results]

            # 左图：盈利概率对比
            colors_prob = ['green' if p >= 50 else 'orange' if p >= 30 else 'red' for p in profit_probs]
            bars1 = ax1.bar(scenario_names, profit_probs, color=colors_prob, alpha=0.7, edgecolor='black')
            ax1.set_ylabel('盈利概率 (%)', fontproperties=font_prop, fontsize=12)
            ax1.set_title('各情景盈利概率对比', fontproperties=font_prop, fontsize=13, fontweight='bold')
            ax1.set_ylim(0, 100)
            ax1.axhline(y=50, color='gray', linestyle='--', alpha=0.5, label='盈亏平衡线')
            ax1.legend(prop=font_prop)

            # 添加数值标注
            for bar in bars1:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height,
                        f'{height:.1f}%', ha='center', va='bottom', fontproperties=font_prop, fontsize=10)

            # 右图：收益率中位数对比
            colors_return = ['green' if r >= 0 else 'red' for r in median_returns]
            bars2 = ax2.bar(scenario_names, median_returns, color=colors_return, alpha=0.7, edgecolor='black')
            ax2.set_ylabel('收益率中位数 (%)', fontproperties=font_prop, fontsize=12)
            ax2.set_title('各情景收益率中位数对比', fontproperties=font_prop, fontsize=13, fontweight='bold')
            ax2.axhline(y=0, color='black', linestyle='-', alpha=0.3)

            # 添加数值标注
            for bar in bars2:
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height,
                        f'{height:+.1f}%', ha='center', va='bottom' if height >= 0 else 'top',
                        fontproperties=font_prop, fontsize=10)

            plt.tight_layout()

            # 保存图表
            scenario_chart_path = os.path.join(IMAGES_DIR, '06_02_scenario_comparison.png')
            plt.savefig(scenario_chart_path, dpi=150, bbox_inches='tight')
            plt.close()

            add_paragraph(document, '图表 6.2: 市场指数与行业情景对比分析')
            add_image(document, scenario_chart_path, width=Inches(6))

            add_paragraph(document, '')
            add_paragraph(document, '📊 情景分析结论：')
            best_scenario = max(index_industry_results, key=lambda x: x['profit_prob'])
            worst_scenario = min(index_industry_results, key=lambda x: x['profit_prob'])
            add_paragraph(document, f'• 最优情景：{best_scenario["scenario"]["name"]}，盈利概率{best_scenario["profit_prob"]:.1f}%，收益率中位数{best_scenario["median_return"]:+.1f}%')
            add_paragraph(document, f'• 最差情景：{worst_scenario["scenario"]["name"]}，盈利概率{worst_scenario["profit_prob"]:.1f}%，收益率中位数{worst_scenario["median_return"]:+.1f}%')
            add_paragraph(document, '• 投资建议：关注市场环境和行业趋势，在乐观环境下可适当提高仓位，悲观环境下需谨慎评估')

        except Exception as e:
            print(f"⚠️ 生成情景对比图表失败: {e}")
            import traceback
            traceback.print_exc()

    # 根据盈利概率和收益率中位数排序并分类
    if comprehensive_results:
        # 按盈利概率和收益率中位数综合排序
        comprehensive_results.sort(key=lambda x: (x['profit_prob'], x['median_return']), reverse=True)

        # 分类
        n = len(comprehensive_results)
        categories = []
        if n >= 3:
            third = n // 3
            categories = ['乐观'] * third + ['中性'] * (third if n >= 2*third else n - 2*third) + ['悲观'] * (n - 2*third)
        else:
            categories = ['中性'] * n

        # ==================== 6.3 基于行业PE分位数的情景分析 ====================
        add_paragraph(document, '')
        add_title(document, '6.3 基于行业PE分位数的情景分析', level=2)

        add_paragraph(document, '本节基于行业PE分位数的估值水平，构建不同估值情景进行蒙特卡洛模拟分析。')
        add_paragraph(document, '通过分析行业PE 75%、50%、25%分位数下的情景，评估不同估值水平对定增收益的影响。')
        add_paragraph(document, '')

        add_paragraph(document, '行业PE分位数计算逻辑：')
        add_paragraph(document, '• 收集行业历史PE数据（通常为5年）')
        add_paragraph(document, '• 计算25%、50%、75%分位数，分别代表低估、中性、高估水平')
        add_paragraph(document, '• 漂移率 = (行业PE分位数 / 个股当前PE - 1)，反映估值回归潜力')
        add_paragraph(document, '• 如果个股PE低于行业PE分位数，估值便宜，未来收益率应该更高')
        add_paragraph(document, '')

        # 筛选行业PE情景
        industry_pe_results = [r for r in comprehensive_results
                              if 'scenario' in r and
                              (r['scenario'].get('name', '').startswith('情景4') or
                               r['scenario'].get('name', '').startswith('情景5') or
                               r['scenario'].get('name', '').startswith('情景6'))]

        if industry_pe_results:
            add_paragraph(document, '行业PE分位数情景参数表：')
            industry_pe_table_data = []
            for result in industry_pe_results:
                scenario_obj = result['scenario']
                industry_pe_table_data.append([
                    scenario_obj['name'],
                    scenario_obj['description'],
                    f"{scenario_obj['volatility']*100:.2f}%",
                    f"{scenario_obj['drift']*100:+.2f}%",
                    f"{scenario_obj['premium_rate']*100:+.0f}%",
                    f"{result['profit_prob']:.1f}%",
                    f"{result['median_return']:+.1f}%"
                ])
            add_table_data(document, ['情景名称', '情景描述', '波动率', '漂移率', '溢价率', '盈利概率', '收益率中位数'], industry_pe_table_data, font_size=10.5)

            add_paragraph(document, '')
            add_paragraph(document, '📊 行业PE分位数情景分析：')
            best_industry_pe = max(industry_pe_results, key=lambda x: x['profit_prob'])
            worst_industry_pe = min(industry_pe_results, key=lambda x: x['profit_prob'])
            add_paragraph(document, f'• 最优情景：{best_industry_pe["scenario"]["name"]}，盈利概率{best_industry_pe["profit_prob"]:.1f}%')
            add_paragraph(document, f'• 最差情景：{worst_industry_pe["scenario"]["name"]}，盈利概率{worst_industry_pe["profit_prob"]:.1f}%')
            add_paragraph(document, '• 投资建议：优先选择行业PE分位数较高时的投资机会，估值安全边际充足')

        # ==================== 6.4 基于个股PE分位数的情景分析 ====================
        add_paragraph(document, '')
        add_title(document, '6.4 基于个股PE分位数的情景分析', level=2)

        add_paragraph(document, '本节基于个股PE分位数的估值水平，构建不同估值情景进行蒙特卡洛模拟分析。')
        add_paragraph(document, '通过分析个股PE 75%、50%、25%分位数下的情景，评估个股历史估值水平对定增收益的影响。')
        add_paragraph(document, '')

        add_paragraph(document, '个股PE分位数计算逻辑：')
        add_paragraph(document, '• 收集个股历史PE数据（通常为5年）')
        add_paragraph(document, '• 计算25%、50%、75%分位数，分别代表历史低估、中性、高估水平')
        add_paragraph(document, '• 漂移率 = (个股PE分位数 / 个股当前PE - 1)，反映估值回归潜力')
        add_paragraph(document, '• 如果当前PE低于历史分位数，估值便宜，未来收益率应该更高')
        add_paragraph(document, '')

        # 筛选个股PE情景
        stock_pe_results = [r for r in comprehensive_results
                           if 'scenario' in r and
                           (r['scenario'].get('name', '').startswith('情景7') or
                            r['scenario'].get('name', '').startswith('情景8') or
                            r['scenario'].get('name', '').startswith('情景9'))]

        if stock_pe_results:
            add_paragraph(document, '个股PE分位数情景参数表：')
            stock_pe_table_data = []
            for result in stock_pe_results:
                scenario_obj = result['scenario']
                stock_pe_table_data.append([
                    scenario_obj['name'],
                    scenario_obj['description'],
                    f"{scenario_obj['volatility']*100:.2f}%",
                    f"{scenario_obj['drift']*100:+.2f}%",
                    f"{scenario_obj['premium_rate']*100:+.0f}%",
                    f"{result['profit_prob']:.1f}%",
                    f"{result['median_return']:+.1f}%"
                ])
            add_table_data(document, ['情景名称', '情景描述', '波动率', '漂移率', '溢价率', '盈利概率', '收益率中位数'], stock_pe_table_data, font_size=10.5)

            add_paragraph(document, '')
            add_paragraph(document, '📊 个股PE分位数情景分析：')
            best_stock_pe = max(stock_pe_results, key=lambda x: x['profit_prob'])
            worst_stock_pe = min(stock_pe_results, key=lambda x: x['profit_prob'])
            add_paragraph(document, f'• 最优情景：{best_stock_pe["scenario"]["name"]}，盈利概率{best_stock_pe["profit_prob"]:.1f}%')
            add_paragraph(document, f'• 最差情景：{worst_stock_pe["scenario"]["name"]}，盈利概率{worst_stock_pe["profit_prob"]:.1f}%')
            add_paragraph(document, '• 投资建议：关注个股历史估值水平，当前PE处于历史低位时投资价值更高')

        # ==================== 6.5 基于DCF估值的情景分析 ====================
        add_paragraph(document, '')
        add_title(document, '6.5 基于DCF估值的情景分析', level=2)

        add_paragraph(document, '本节基于DCF绝对估值方法，构建不同估值情景进行蒙特卡洛模拟分析。')
        add_paragraph(document, '通过分析DCF内在价值与当前价格的比值，评估公司内在价值对定增收益的影响。')
        add_paragraph(document, '')

        add_paragraph(document, 'DCF估值方法说明：')
        add_paragraph(document, '• DCF（现金流折现模型）通过预测未来自由现金流并折现计算内在价值')
        add_paragraph(document, '• 漂移率 = (DCF内在价值 / 当前价格 - 1)，反映内在价值与市场价格的偏离')
        add_paragraph(document, '• 如果DCF内在价值高于当前价格，股票被低估，未来收益率应该为正')
        add_paragraph(document, '• DCF估值提供绝对价值判断，独立于市场相对估值')
        add_paragraph(document, '')

        # 筛选DCF情景
        dcf_results = [r for r in comprehensive_results
                      if 'scenario' in r and
                      (r['scenario'].get('name', '').startswith('情景10') or
                       r['scenario'].get('name', '').startswith('情景11') or
                       r['scenario'].get('name', '').startswith('情景12'))]

        if dcf_results:
            add_paragraph(document, 'DCF估值情景参数表：')
            dcf_table_data = []
            for result in dcf_results:
                scenario_obj = result['scenario']
                dcf_table_data.append([
                    scenario_obj['name'],
                    scenario_obj['description'],
                    f"{scenario_obj['volatility']*100:.2f}%",
                    f"{scenario_obj['drift']*100:+.2f}%",
                    f"{scenario_obj['premium_rate']*100:+.0f}%",
                    f"{result['profit_prob']:.1f}%",
                    f"{result['median_return']:+.1f}%"
                ])
            add_table_data(document, ['情景名称', '情景描述', '波动率', '漂移率', '溢价率', '盈利概率', '收益率中位数'], dcf_table_data, font_size=10.5)

            add_paragraph(document, '')
            add_paragraph(document, '📊 DCF估值情景分析：')
            best_dcf = max(dcf_results, key=lambda x: x['profit_prob'])
            worst_dcf = min(dcf_results, key=lambda x: x['profit_prob'])
            add_paragraph(document, f'• 最优情景：{best_dcf["scenario"]["name"]}，盈利概率{best_dcf["profit_prob"]:.1f}%')
            add_paragraph(document, f'• 最差情景：{worst_dcf["scenario"]["name"]}，盈利概率{worst_dcf["profit_prob"]:.1f}%')
            add_paragraph(document, '• 投资建议：DCF内在价值提供安全边际，优先选择DCF估值高于市场价格的项目')

        # ==================== 6.6 情景综合分析汇总表 ====================
        add_paragraph(document, '')
        add_title(document, '6.6 情景综合分析汇总表', level=2)

        add_paragraph(document, '本节汇总所有情景的蒙特卡洛模拟结果，提供综合对比分析。')
        add_paragraph(document, '通过汇总表可以全面评估不同情景下的风险收益特征，为投资决策提供参考。')
        add_paragraph(document, '')

        # 生成表格数据
        comprehensive_table_data = []
        for i, result in enumerate(comprehensive_results):
            scenario_obj = result['scenario']
            category = categories[i] if i < len(categories) else '中性'

            comprehensive_table_data.append([
                scenario_obj['description'],  # 情景描述
                f"{scenario_obj['volatility']*100:.2f}%",
                f"{scenario_obj['drift']*100:+.2f}%",
                f"{scenario_obj['premium_rate']*100:+.0f}%",
                f"{result['issue_price']:.2f}",
                f"{result['profit_prob']:.1f}%",
                f"{result['median_return']:+.1f}%",
                scenario_obj['name']  # 情景类型
            ])

        comprehensive_headers = ['情景描述', '波动率', '漂移率', '溢价率', '发行价(元)', '盈利概率(%)', '收益率中位数(%)', '情景类型']
        add_table_data(document, comprehensive_headers, comprehensive_table_data, font_size=10.5)  # 五号字体

        # 添加说明
        add_paragraph(document, '')
        add_paragraph(document, '情景说明：')
        add_paragraph(document, '• 当前情景：标的股票120日窗口真实数据，反映项目实际风险收益特征')
        add_paragraph(document, '• 情景1-3：基于行业120日窗口历史数据的典型情景（乐观/中性/悲观）')
        if pe_scenarios:
            add_paragraph(document, '• 情景4-6：基于行业PE分位数的估值情景，反映不同估值水平下的风险收益')
            add_paragraph(document, '• 情景7-9：基于个股PE分位数的估值情景，反映个股历史估值水平的影响')
            if intrinsic_value is not None:
                add_paragraph(document, '• 情景10-12：基于DCF绝对估值的情景，反映公司内在价值对收益的影响')
        add_paragraph(document, '')
        add_paragraph(document, '📊 评级说明：')
        add_paragraph(document, '• 乐观：盈利概率和收益率中位数均较高，投资价值突出')
        add_paragraph(document, '• 中性：盈利概率和收益率中位数适中，风险可控')
        add_paragraph(document, '• 悲观：盈利概率或收益率中位数较低，风险较高')
        add_paragraph(document, '')
        add_paragraph(document, '投资建议：')
        add_paragraph(document, '• 优先选择乐观评级情景，安全边际充足')
        add_paragraph(document, '• 当前项目定位在"当前情景"，需结合实际参数评估')
        add_paragraph(document, '• 不同情景反映市场环境变化，建议根据风险偏好选择')

    # 生成多维度情景图表（补充分析）
    # 注：波动率×折价率热力图已在6.3章节展示，此处不再重复
    try:
        multi_dim_chart_paths = generate_multi_dimension_scenario_charts_split(
            project_params['current_price'], ma20, risk_params['volatility'],
            risk_params['drift'], project_params['lockup_period'], IMAGES_DIR)

        # 该函数只返回1个图表（3D分析图）
        if len(multi_dim_chart_paths) > 0:
            add_paragraph(document, '图表 6.6: 多维度情景分析（波动率 × 时间窗口 × 折扣率）')
            add_image(document, multi_dim_chart_paths[0], width=Inches(6.5))
            add_paragraph(document, '')

    except Exception as e:
        print(f"⚠️ 生成多维度图表失败: {e}")
        import traceback
        traceback.print_exc()

    add_section_break(document)

    # 保存all_scenarios到context供附件使用
    if 'all_scenarios' in locals() and len(all_scenarios) > 0:
        context['results']['all_scenarios'] = all_scenarios
        print(f"✅ 已保存{len(all_scenarios)}个情景到context，供附件使用")

    return context
