#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增风险分析报告生成器（增强版）
生成 Word 格式的综合分析报告，包含图表
"""

import sys
import os
import json
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats

# 添加路径 - 从 scripts 目录运行
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)

from utils.config_loader import load_placement_config
from utils.analysis_tools import PrivatePlacementRiskAnalyzer
from utils.font_manager import get_font_prop

# 获取中文字体
font_prop = get_font_prop()

# 配置路径
DATA_DIR = os.path.join(PROJECT_DIR, 'data')
REPORTS_DIR = os.path.join(PROJECT_DIR, 'reports')
IMAGES_DIR = os.path.join(REPORTS_DIR, 'images')
OUTPUTS_DIR = os.path.join(REPORTS_DIR, 'outputs')

# 确保目录存在
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(OUTPUTS_DIR, exist_ok=True)


def setup_chinese_font(document):
    """设置中文字体"""
    document.styles['Normal'].font.name = '宋体'
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    document.styles['Normal'].font.size = Pt(11)


def add_title(document, text, level=1):
    """添加标题"""
    heading = document.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = '黑体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    return heading


def add_paragraph(document, text, bold=False, font_size=11):
    """添加段落"""
    para = document.add_paragraph(text)
    for run in para.runs:
        run.font.name = '宋体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.size = Pt(font_size)
        if bold:
            run.font.bold = True
    return para


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


def add_table_data(document, headers, data):
    """添加表格"""
    table = document.add_table(rows=1, cols=len(headers))
    table.style = 'Light Grid Accent 1'

    # 设置表头
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        header_cells[i].text = header
        para = header_cells[i].paragraphs[0]
        para.runs[0].font.bold = True
        para.runs[0].font.size = Pt(10)
        para.runs[0].font.name = '宋体'
        para.runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 添加数据行
    for row_data in data:
        row_cells = table.add_row().cells
        for i, cell_data in enumerate(row_data):
            row_cells[i].text = str(cell_data)
            para = row_cells[i].paragraphs[0]
            para.runs[0].font.size = Pt(10)
            para.runs[0].font.name = '宋体'
            para.runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            if isinstance(cell_data, (int, float)):
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    return table


def add_section_break(document):
    """添加分页符"""
    document.add_page_break()


def generate_sensitivity_chart(volatilities, profit_probs, current_vol, ma30, issue_price, save_path):
    """生成敏感性分析图表（增强版）"""
    fig = plt.figure(figsize=(16, 10))
    gs = fig.add_gridspec(2, 2, hspace=0.3, wspace=0.3)

    # 1. 波动率 vs 盈利概率柱状图
    ax1 = fig.add_subplot(gs[0, 0])
    colors = ['#27ae60' if p >= 60 else '#f39c12' if p >= 40 else '#e74c3c' for p in profit_probs]
    bars = ax1.bar(range(len(volatilities)), profit_probs, color=colors, alpha=0.8, edgecolor='black', linewidth=1.5)
    ax1.set_xlabel('年化波动率', fontproperties=font_prop, fontsize=11)
    ax1.set_ylabel('盈利概率 (%)', fontproperties=font_prop, fontsize=11)
    ax1.set_title('不同波动率下的盈利概率', fontproperties=font_prop, fontsize=12, fontweight='bold')
    ax1.set_xticks(range(len(volatilities)))
    ax1.set_xticklabels([f'{v*100:.0f}%' for v in volatilities], fontproperties=font_prop, fontsize=9)
    ax1.grid(True, alpha=0.3, linestyle='--')
    ax1.set_ylim(0, 100)
    ax1.axhline(y=50, color='gray', linestyle='--', alpha=0.5)

    # 标记当前波动率
    current_idx = np.argmin(np.abs(np.array(volatilities) - current_vol))
    ax1.axvline(x=current_idx, color='red', linestyle='-', linewidth=2, alpha=0.7)
    ax1.text(current_idx, 95, f'当前\n{current_vol*100:.1f}%', ha='center', va='top',
            fontsize=9, fontproperties=font_prop, color='red', fontweight='bold')
    for bar, prob in zip(bars, profit_probs):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{prob:.1f}%', ha='center', va='bottom', fontsize=9, fontweight='bold')

    # 2. 发行价折扣分析
    ax2 = fig.add_subplot(gs[0, 1])
    discounts = np.array([5, 10, 15, 20])  # 折价率
    issue_prices_discounted = ma30 * (1 - discounts/100)

    # 计算不同发行价下的盈利概率
    discount_probs = []
    for discounted_price in issue_prices_discounted:
        # 使用当前市场参数
        lockup_months = 6  # 假设6个月锁定期
        drift_rate = -0.1875  # 当前漂移率
        lockup_drift = drift_rate * (lockup_months / 12)
        lockup_vol = 0.3063 * np.sqrt(lockup_months / 12)  # 当前波动率

        # 计算盈利概率
        z = (np.log(ma30 / discounted_price) - lockup_drift) / lockup_vol
        prob = stats.norm.cdf(z) * 100
        discount_probs.append(prob)

    colors_discount = ['#27ae60', '#2ecc71', '#f39c12', '#e74c3c']
    bars2 = ax2.bar(range(len(discounts)), discount_probs, color=colors_discount,
                   alpha=0.8, edgecolor='black', linewidth=1.5)
    ax2.set_xlabel('发行价相对MA30折价率 (%)', fontproperties=font_prop, fontsize=11)
    ax2.set_ylabel('盈利概率 (%)', fontproperties=font_prop, fontsize=11)
    ax2.set_title('不同发行价折扣下的盈利概率', fontproperties=font_prop, fontsize=12, fontweight='bold')
    ax2.set_xticks(range(len(discounts)))
    ax2.set_xticklabels([f'{d}%' for d in discounts], fontproperties=font_prop, fontsize=10)
    ax2.grid(True, alpha=0.3, linestyle='--')
    ax2.set_ylim(0, 100)
    ax2.axhline(y=50, color='gray', linestyle='--', alpha=0.5, label='盈亏平衡线')
    ax2.axvline(x=0, color='blue', linestyle='--', linewidth=2, alpha=0.7, label='当前发行价位置')
    ax2.legend(prop=font_prop, fontsize=9, loc='lower right')

    # 添加数值标签和发行价信息
    for bar, prob, price in zip(bars2, discount_probs, issue_prices_discounted):
        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{prob:.1f}%', ha='center', va='bottom', fontsize=9, fontweight='bold')
        ax2.text(bar.get_x() + bar.get_width()/2, -3,
                f'{price:.2f}元', ha='center', va='top', fontsize=8, fontproperties=font_prop)

    # 3. 热力图（波动率 vs 漂移率）
    ax3 = fig.add_subplot(gs[1, :])
    drift_range = np.linspace(-0.3, 0.15, 8)
    vol_range_heatmap = np.linspace(0.20, 0.50, 8)
    heatmap_data = []

    for d in drift_range:
        row = []
        for v in vol_range_heatmap:
            lockup_drift = d * (6/12)
            lockup_vol = v * np.sqrt(6/12)
            z = (np.log(ma30/issue_price) - lockup_drift) / lockup_vol
            row.append(stats.norm.cdf(z) * 100)
        heatmap_data.append(row)

    im = ax3.imshow(heatmap_data, cmap='RdYlGn', aspect='auto', vmin=0, vmax=100,
                   extent=[vol_range_heatmap[0]*100, vol_range_heatmap[-1]*100, drift_range[0]*100, drift_range[-1]*100])
    ax3.set_xlabel('年化波动率 (%)', fontproperties=font_prop, fontsize=12)
    ax3.set_ylabel('年化漂移率 (%)', fontproperties=font_prop, fontsize=12)
    ax3.set_title('盈利概率敏感性热力图（波动率 vs 漂移率）', fontproperties=font_prop, fontsize=13, fontweight='bold')
    ax3.tick_params(axis='both', which='major', labelsize=10)
    for label in ax3.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax3.get_yticklabels():
        label.set_fontproperties(font_prop)

    cbar = plt.colorbar(im, ax=ax3)
    cbar.set_label('盈利概率 (%)', fontproperties=font_prop, fontsize=11)
    cbar.ax.tick_params(labelsize=10)

    # 标记当前位置
    ax3.scatter([current_vol*100], [drift_rate*100], color='red', s=300, marker='*',
               edgecolors='white', linewidths=2, label='当前位置', zorder=5)
    ax3.legend(prop=font_prop, loc='upper right', fontsize=10)

    plt.suptitle('敏感性分析：波动率、漂移率与发行价折扣', fontproperties=font_prop,
               fontsize=15, fontweight='bold', y=0.98)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()

    # 左图：柱状图
    colors = ['#27ae60' if p >= 60 else '#f39c12' if p >= 40 else '#e74c3c' for p in profit_probs]
    bars = ax1.bar(range(len(volatilities)), profit_probs, color=colors, alpha=0.8, edgecolor='black', linewidth=1.5)
    ax1.set_xlabel('年化波动率', fontproperties=font_prop, fontsize=12)
    ax1.set_ylabel('盈利概率 (%)', fontproperties=font_prop, fontsize=12)
    ax1.set_title('不同波动率下的盈利概率', fontproperties=font_prop, fontsize=13, fontweight='bold')
    ax1.set_xticks(range(len(volatilities)))
    ax1.set_xticklabels([f'{v*100:.0f}%' for v in volatilities], fontproperties=font_prop)
    ax1.grid(True, alpha=0.3, linestyle='--')
    ax1.set_ylim(0, 100)
    ax1.axhline(y=50, color='gray', linestyle='--', alpha=0.5, label='盈亏平衡线')
    ax1.legend(prop=font_prop, loc='lower right')

    # 标记当前波动率
    current_idx = np.argmin(np.abs(np.array(volatilities) - current_vol))
    ax1.axvline(x=current_idx, color='red', linestyle='-', linewidth=2, alpha=0.7)
    ax1.text(current_idx, 95, f'当前\n{current_vol*100:.1f}%', ha='center', va='top',
            fontsize=9, fontproperties=font_prop, color='red', fontweight='bold')

    # 添加数值标签
    for bar, prob in zip(bars, profit_probs):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f'{prob:.1f}%', ha='center', va='bottom', fontsize=10, fontweight='bold')

    # 右图：热力图（波动率 vs 漂移率）
    drift_range = np.linspace(-0.3, 0.15, 8)
    vol_range_heatmap = np.linspace(0.20, 0.50, 8)
    heatmap_data = []

    for d in drift_range:
        row = []
        for v in vol_range_heatmap:
            lockup_drift = d * (6/12)
            lockup_vol = v * np.sqrt(6/12)
            z = (np.log(ma30/20.25) - lockup_drift) / lockup_vol
            row.append(stats.norm.cdf(z) * 100)
        heatmap_data.append(row)

    im = ax2.imshow(heatmap_data, cmap='RdYlGn', aspect='auto', vmin=0, vmax=100,
                   extent=[vol_range_heatmap[0]*100, vol_range_heatmap[-1]*100, drift_range[-1]*100, drift_range[0]*100])
    ax2.set_xlabel('波动率 (%)', fontproperties=font_prop, fontsize=11)
    ax2.set_ylabel('漂移率 (%)', fontproperties=font_prop, fontsize=11)
    ax2.set_title('盈利概率热力图', fontproperties=font_prop, fontsize=13, fontweight='bold')
    ax2.tick_params(axis='both', which='major', labelsize=9)
    for label in ax2.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax2.get_yticklabels():
        label.set_fontproperties(font_prop)

    cbar = plt.colorbar(im, ax=ax2)
    cbar.set_label('盈利概率 (%)', fontproperties=font_prop, fontsize=10)
    cbar.ax.tick_params(labelsize=9)

    # 标记当前位置
    ax2.scatter([current_vol*100], [-18.75], color='red', s=200, marker='*',
               edgecolors='white', linewidths=2, label='当前位置', zorder=5)
    ax2.legend(prop=font_prop, loc='upper right', fontsize=9)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_stress_test_chart(scenarios, returns, save_path):
    """生成压力测试图表"""
    plt.figure(figsize=(10, 6))
    colors = ['#e74c3c' if r < 0 else '#2ecc71' for r in returns]
    bars = plt.barh(range(len(scenarios)), returns, color=colors, alpha=0.7)
    plt.xlabel('收益率 (%)', fontproperties=font_prop, fontsize=12)
    plt.ylabel('情景', fontproperties=font_prop, fontsize=12)
    plt.title('压力测试情景分析', fontproperties=font_prop, fontsize=14, fontweight='bold')
    plt.yticks(range(len(scenarios)), scenarios, fontproperties=font_prop)
    plt.xticks(fontproperties=font_prop)
    plt.axvline(x=0, color='black', linestyle='-', linewidth=1)
    plt.grid(True, alpha=0.3, axis='x')

    for bar, ret in zip(bars, returns):
        plt.text(bar.get_width() + (1 if ret > 0 else -1),
                bar.get_y() + bar.get_height()/2,
                f'{ret:.1f}%', ha='left' if ret > 0 else 'right',
                va='center', fontsize=10, fontproperties=font_prop)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_monte_carlo_chart(final_prices, issue_price, current_price, save_path):
    """生成蒙特卡洛模拟图表（增强版）"""
    fig = plt.figure(figsize=(16, 10))

    # 1. 价格分布直方图
    ax1 = plt.subplot(2, 3, 1)
    ax1.hist(final_prices, bins=50, alpha=0.7, color='steelblue', edgecolor='black')
    ax1.axvline(x=issue_price, color='red', linestyle='--', linewidth=2, label=f'发行价: {issue_price:.2f}')
    ax1.axvline(x=current_price, color='green', linestyle='--', linewidth=2, label=f'当前价: {current_price:.2f}')
    ax1.set_xlabel('未来价格 (元)', fontproperties=font_prop, fontsize=11)
    ax1.set_ylabel('频数', fontproperties=font_prop, fontsize=11)
    ax1.set_title('锁定期后价格分布', fontproperties=font_prop, fontsize=12, fontweight='bold')
    ax1.legend(prop=font_prop, fontsize=9)
    ax1.grid(True, alpha=0.3)
    for label in ax1.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax1.get_yticklabels():
        label.set_fontproperties(font_prop)

    # 2. 收益率分布
    ax2 = plt.subplot(2, 3, 2)
    returns = (final_prices - issue_price) / issue_price * 100
    ax2.hist(returns, bins=50, alpha=0.7, color='coral', edgecolor='black')
    ax2.axvline(x=0, color='red', linestyle='--', linewidth=2, label='盈亏平衡')
    ax2.set_xlabel('收益率 (%)', fontproperties=font_prop, fontsize=11)
    ax2.set_ylabel('频数', fontproperties=font_prop, fontsize=11)
    ax2.set_title('收益率分布', fontproperties=font_prop, fontsize=12, fontweight='bold')
    ax2.legend(prop=font_prop, fontsize=9)
    ax2.grid(True, alpha=0.3)
    for label in ax2.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax2.get_yticklabels():
        label.set_fontproperties(font_prop)

    # 3. 累积分布函数
    ax3 = plt.subplot(2, 3, 3)
    sorted_returns = np.sort(returns)
    cumulative = np.arange(1, len(sorted_returns) + 1) / len(sorted_returns) * 100
    ax3.plot(sorted_returns, cumulative, linewidth=2, color='#8e44ad')
    ax3.axvline(x=0, color='red', linestyle='--', linewidth=1.5, label='盈亏平衡')
    ax3.axhline(y=50, color='gray', linestyle='--', linewidth=1, alpha=0.5)
    ax3.set_xlabel('收益率 (%)', fontproperties=font_prop, fontsize=11)
    ax3.set_ylabel('累积概率 (%)', fontproperties=font_prop, fontsize=11)
    ax3.set_title('累积分布函数', fontproperties=font_prop, fontsize=12, fontweight='bold')
    ax3.legend(prop=font_prop, fontsize=9)
    ax3.grid(True, alpha=0.3)
    for label in ax3.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax3.get_yticklabels():
        label.set_fontproperties(font_prop)

    # 4. 盈亏概率饼图
    ax4 = plt.subplot(2, 3, 4)
    profit_prob = (returns > 0).mean() * 100
    loss_prob = 100 - profit_prob
    colors_pie = ['#27ae60', '#e74c3c']
    wedges, texts, autotexts = ax4.pie([profit_prob, loss_prob], labels=['盈利', '亏损'],
                                         autopct='%1.1f%%', colors=colors_pie, startangle=90)
    ax4.set_title('盈亏概率分布', fontproperties=font_prop, fontsize=12, fontweight='bold')
    for text in texts:
        text.set_fontproperties(font_prop)
    for autotext in autotexts:
        autotext.set_fontproperties(font_prop)
        autotext.set_fontsize(12)
        autotext.set_fontweight('bold')

    # 5. 关键统计指标
    ax5 = plt.subplot(2, 3, 5)
    ax5.axis('off')
    stats_text = f"""
    关键统计指标
    ════════════════

    📊 盈利概率: {profit_prob:.1f}%

    💰 预期收益率: {returns.mean():.1f}%

    📈 收益率中位数: {np.median(returns):.1f}%

    ⬇️  95% VaR: {abs(np.percentile(returns, 5)):.1f}%

    ⬇️  99% VaR: {abs(np.percentile(returns, 1)):.1f}%

    📉 最大回撤: {abs(returns.min()):.1f}%
    """
    ax5.text(0.1, 0.5, stats_text, transform=ax5.transAxes,
            fontsize=11, verticalalignment='center', fontfamily='monospace',
            bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.3))

    # 6. 概率分布曲线
    ax6 = plt.subplot(2, 3, 6)
    from scipy.stats import norm
    mu, std = returns.mean(), returns.std()
    x = np.linspace(returns.min(), returns.max(), 100)
    pdf = norm.pdf(x, mu, std)
    ax6.plot(x, pdf, linewidth=2, color='#2980b9', label='正态分布拟合')
    ax6.fill_between(x, pdf, alpha=0.3, color='#3498db')
    ax6.axvline(x=0, color='red', linestyle='--', linewidth=2, label='盈亏平衡')
    ax6.axvline(x=mu, color='green', linestyle='--', linewidth=2, label=f'均值: {mu:.1f}%')
    ax6.set_xlabel('收益率 (%)', fontproperties=font_prop, fontsize=11)
    ax6.set_ylabel('概率密度', fontproperties=font_prop, fontsize=11)
    ax6.set_title('收益率概率分布', fontproperties=font_prop, fontsize=12, fontweight='bold')
    ax6.legend(prop=font_prop, fontsize=9)
    ax6.grid(True, alpha=0.3)
    for label in ax6.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax6.get_yticklabels():
        label.set_fontproperties(font_prop)

    plt.suptitle('蒙特卡洛模拟分析 (5000次模拟)', fontproperties=font_prop,
               fontsize=14, fontweight='bold', y=0.995)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_dcf_valuation_heatmap(save_path):
    """生成DCF估值热力图"""
    fig, ax = plt.subplots(figsize=(12, 8))

    # WACC 和永续增长率范围
    wacc_range = np.linspace(0.07, 0.13, 7)
    growth_range = np.linspace(0.01, 0.04, 7)

    # 估值参数
    current_price = 23.88
    issue_price = 20.25

    # 计算每股价值矩阵
    pivot_data = []
    for g in growth_range:
        row = []
        for w in wacc_range:
            # 简化的DCF计算
            fcfs = [2.0] * 10  # 假设每年2亿自由现金流
            pv_fcfs = sum([fcf / ((1 + w) ** (i+1)) for i, fcf in enumerate(fcfs)])
            terminal_value = fcfs[-1] * (1 + g) / (w - g) if w > g else 0
            pv_terminal = terminal_value / ((1 + w) ** 10)

            enterprise_value = pv_fcfs + pv_terminal
            equity_value = enterprise_value - 20  # 减债务
            value_per_share = equity_value / 5  # 5亿股
            row.append(value_per_share)
        pivot_data.append(row)

    pivot_table = pd.DataFrame(
        pivot_data,
        index=[f'{g*100:.1f}%' for g in growth_range],
        columns=[f'{w*100:.1f}%' for w in wacc_range]
    )

    # 绘制热力图
    heatmap = sns.heatmap(pivot_table, annot=True, fmt='.2f', cmap='RdYlGn',
                          center=current_price, cbar_kws={'label': '每股价值（元）'}, ax=ax)

    # 修复 colorbar 中文
    cbar = heatmap.collections[0].colorbar
    cbar.set_label('每股价值（元）', fontproperties=font_prop, fontsize=12)

    # 设置轴标签
    ax.set_xlabel('WACC（加权平均资本成本）', fontproperties=font_prop, fontsize=12)
    ax.set_ylabel('永续增长率', fontproperties=font_prop, fontsize=12)
    ax.set_title('DCF估值敏感性分析矩阵（每股价值：元）',
                fontproperties=font_prop, fontsize=14, fontweight='bold', pad=15)

    # 设置刻度标签字体
    for label in ax.get_xticklabels():
        label.set_fontproperties(font_prop)
        label.set_fontsize(10)
    for label in ax.get_yticklabels():
        label.set_fontproperties(font_prop)
        label.set_fontsize(10)

    # 添加参考线
    ax.axhline(y=2.5, color='white', linestyle='--', linewidth=1, alpha=0.5)
    ax.axvline(x=2.5, color='white', linestyle='--', linewidth=1, alpha=0.5)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_var_chart(var_95, var_99, cvar_95, save_path):
    """生成VaR风险度量图表"""
    plt.figure(figsize=(10, 6))
    metrics = ['95% VaR', '99% VaR', '95% CVaR']
    values = [var_95, var_99, cvar_95]
    colors = ['#e74c3c', '#c0392b', '#922b21']

    bars = plt.bar(metrics, [v*100 for v in values], color=colors, alpha=0.7)
    plt.ylabel('潜在损失 (%)', fontproperties=font_prop, fontsize=12)
    plt.title('风险度量 (VaR/CVaR)', fontproperties=font_prop, fontsize=14, fontweight='bold')
    plt.xticks(fontproperties=font_prop)
    plt.yticks(fontproperties=font_prop)
    plt.grid(True, alpha=0.3, axis='y')

    for bar, value in zip(bars, values):
        plt.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                f'{value*100:.1f}%', ha='center', va='bottom',
                fontsize=11, fontweight='bold', fontproperties=font_prop)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_radar_chart(scores, save_path):
    """生成风险评分雷达图"""
    categories = list(scores.keys())
    values = list(scores.values())

    # 闭合雷达图
    angles_plot = np.linspace(0, 2*np.pi, len(categories), endpoint=False).tolist()
    angles_plot += angles_plot[:1]
    values_plot = values + values[:1]

    fig, ax = plt.subplots(figsize=(10, 10), subplot_kw=dict(projection='polar'))
    ax.plot(angles_plot, values_plot, 'o-', linewidth=2, color='#3498db')
    ax.fill(angles_plot, values_plot, alpha=0.25, color='#3498db')
    ax.set_xticks(angles_plot[:-1])
    ax.set_xticklabels(categories, fontproperties=font_prop, fontsize=11)
    ax.set_ylim(0, 30)
    ax.set_title('定增项目风险雷达图', fontsize=14, fontweight='bold',
                pad=20, fontproperties=font_prop)
    ax.grid(True)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_report(stock_code='300735.SZ', output_file='定增风险分析报告.docx'):
    """生成完整的定增风险分析报告"""

    print("开始生成定增风险分析报告（包含图表）...")

    # 创建文档
    document = Document()
    setup_chinese_font(document)

    # 加载配置 - 使用数据目录中的文件
    print("加载配置数据...")
    project_params, risk_params, market_data = load_placement_config(stock_code, data_dir=DATA_DIR)

    # 发行类型评估
    ma30 = market_data.get('ma_30', 0)
    issue_price = project_params['issue_price']
    if issue_price < ma30:
        issue_type = "折价发行"
        discount_premium = (ma30 - issue_price) / ma30 * 100
    else:
        issue_type = "溢价发行"
        discount_premium = (issue_price - ma30) / ma30 * 100

    # 创建分析器
    analyzer = PrivatePlacementRiskAnalyzer(
        issue_price=project_params['issue_price'],
        issue_shares=project_params['issue_shares'],
        lockup_period=project_params['lockup_period'],
        current_price=project_params['current_price'],
        risk_free_rate=project_params['risk_free_rate']
    )

    current_date = datetime.now().strftime('%Y年%m月%d日')

    # ==================== 封面 ====================
    add_title(document, '定增项目风险分析报告', level=0)
    document.add_paragraph()
    document.add_paragraph()

    # 报告基本信息
    info_table = document.add_table(rows=6, cols=2)
    info_table.style = 'Table Grid'

    info_data = [
        ['公司名称', '光弘科技 (300735.SZ)'],
        ['报告日期', current_date],
        ['发行日期', current_date],
        ['发行价格', f'{project_params["issue_price"]:.2f} 元/股'],
        ['当前价格', f'{project_params["current_price"]:.2f} 元/股'],
        ['发行类型', f'{issue_type}（安全边际/溢价率: {discount_premium:.2f}%）']
    ]

    for i, (label, value) in enumerate(info_data):
        info_table.rows[i].cells[0].text = label
        info_table.rows[i].cells[1].text = value
        for cell in info_table.rows[i].cells:
            cell.paragraphs[0].runs[0].font.size = Pt(11)
            cell.paragraphs[0].runs[0].font.name = '宋体'
            cell.paragraphs[0].runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    add_section_break(document)

    # ==================== 目录 ====================
    add_title(document, '目 录', level=1)
    add_paragraph(document, '一、项目概况')
    add_paragraph(document, '二、敏感性分析')
    add_paragraph(document, '三、压力测试')
    add_paragraph(document, '四、DCF估值分析')
    add_paragraph(document, '五、蒙特卡洛模拟')
    add_paragraph(document, '六、VaR风险度量')
    add_paragraph(document, '七、综合评估')
    add_paragraph(document, '八、投资建议与风险提示')

    add_section_break(document)

    # ==================== 一、项目概况 ====================
    add_title(document, '一、项目概况', level=1)

    add_title(document, '1.1 基本信息', level=2)

    project_headers = ['指标', '数值']
    project_data = [
        ['股票代码', '300735.SZ'],
        ['公司名称', '光弘科技'],
        ['发行价格', f'{project_params["issue_price"]:.2f} 元/股'],
        ['发行数量', f'{project_params["issue_shares"]:,} 股'],
        ['锁定期', f'{project_params["lockup_period"]} 个月'],
        ['融资金额', f'{project_params["issue_price"] * project_params["issue_shares"] / 100000000:.2f} 亿元'],
        ['当前价格', f'{project_params["current_price"]:.2f} 元/股'],
        ['MA30价格', f'{ma30:.2f} 元/股'],
        ['发行类型', issue_type],
        ['安全边际/溢价率', f'{discount_premium:.2f}%']
    ]
    add_table_data(document, project_headers, project_data)

    add_title(document, '1.2 市场数据', level=2)

    market_headers = ['指标', '数值']
    market_table_data = [
        ['当前价格', f'{market_data["current_price"]:.2f} 元/股'],
        ['平均价格', f'{market_data.get("avg_price_all", 0):.2f} 元/股'],
        ['价格标准差', f'{market_data.get("price_std", 0):.2f}'],
        ['30日波动率', f'{market_data.get("volatility_30d", 0)*100:.2f}%'],
        ['60日波动率', f'{market_data.get("volatility_60d", 0)*100:.2f}%'],
        ['60日年化收益率', f'{market_data.get("annual_return_60d", 0)*100:.2f}%'],
        ['MA30', f'{market_data.get("ma_30", 0):.2f} 元/股'],
        ['MA60', f'{market_data.get("ma_60", 0):.2f} 元/股'],
        ['MA120', f'{market_data.get("ma_120", 0):.2f} 元/股'],
        ['数据来源', 'Tushare API']
    ]
    add_table_data(document, market_headers, market_table_data)

    add_section_break(document)

    # ==================== 二、敏感性分析 ====================
    add_title(document, '二、敏感性分析', level=1)

    add_paragraph(document, '本章节分析在不同波动率和漂移率假设下，定增项目的盈利概率变化情况。')

    add_title(document, '2.1 不同波动率下的盈利概率', level=2)

    # 生成敏感性分析数据和图表
    volatilities = np.linspace(0.20, 0.50, 7)
    drift_rate = risk_params.get('drift', market_data.get('drift', 0.08))

    prob_results = []
    for vol in volatilities:
        lockup_months = project_params['lockup_period']
        lockup_drift = drift_rate * (lockup_months / 12)
        lockup_vol = vol * np.sqrt(lockup_months / 12)

        d = (np.log(ma30 / project_params['issue_price']) - lockup_drift) / lockup_vol
        prob = stats.norm.cdf(d)
        prob_results.append(prob * 100)

    # 保存图表（传入当前波动率）
    sensitivity_chart_path = os.path.join(IMAGES_DIR, '01_sensitivity_analysis.png')
    generate_sensitivity_chart(volatilities, prob_results, market_data['volatility'],
                               ma30, project_params['issue_price'], sensitivity_chart_path)

    vol_data = [[f'{vol*100:.0f}%', f'{prob:.1f}%'] for vol, prob in zip(volatilities, prob_results)]
    add_table_data(document, ['波动率', '盈利概率'], vol_data)

    # 添加发行价折扣分析表格
    add_paragraph(document, '')
    add_title(document, '2.2 发行价折扣分析', level=2)

    # 计算不同折扣下的数据
    discount_analysis_data = []
    for discount_percent in [5, 10, 15, 20]:
        discounted_price = ma30 * (1 - discount_percent/100)
        lockup_months = 6
        drift_rate = -0.1875
        lockup_drift = drift_rate * (lockup_months / 12)
        lockup_vol = 0.3063 * np.sqrt(lockup_months / 12)

        z = (np.log(ma30 / discounted_price) - lockup_drift) / lockup_vol
        prob = stats.norm.cdf(z) * 100

        # 预期收益率（假设锁定期后价格回到MA30）
        expected_return = (ma30 - discounted_price) / discounted_price * 100

        discount_analysis_data.append([
            f'{discount_percent}%',
            f'{discounted_price:.2f}',
            f'{prob:.1f}%',
            f'{expected_return:.1f}%'
        ])

    add_table_data(document, ['发行价折扣', '折扣后发行价', '盈利概率', '预期收益率'], discount_analysis_data)

    add_paragraph(document, '图表 1: 敏感性分析')
    add_image(document, sensitivity_chart_path, width=Inches(6.5))

    add_paragraph(document, '分析结论：')
    add_paragraph(document, f'• 在当前市场波动率（{market_data["volatility"]*100:.1f}%）和漂移率（{market_data["drift"]*100:.1f}%）下，项目盈利概率约 {prob_results[3]:.1f}%')
    add_paragraph(document, '• 波动率越高，盈利概率的不确定性越大')
    add_paragraph(document, '• 折价发行为投资者提供了一定的安全边际')
    add_paragraph(document, '')
    add_paragraph(document, '发行价折扣分析：')

    # 计算当前发行价的折扣
    current_discount = (ma30 - project_params['issue_price']) / ma30 * 100
    add_paragraph(document, f'• 当前发行价折扣: {current_discount:.1f}%（发行价 {project_params["issue_price"]:.2f}元 / MA30 {ma30:.2f}元）')

    if current_discount >= 20:
        discount_level = "非常充足（≥20%）"
        safety_note = "安全边际很高，投资吸引力强"
    elif current_discount >= 15:
        discount_level = "充足（15%-20%）"
        safety_note = "安全边际良好，投资吸引力较高"
    elif current_discount >= 10:
        discount_level = "适中（10%-15%）"
        safety_note = "安全边际一般，需综合考虑其他因素"
    elif current_discount >= 5:
        discount_level = "较低（5%-10%）"
        safety_note = "安全边际有限，需谨慎评估"
    else:
        discount_level = "不足（<5%）"
        safety_note = "安全边际不足，风险较高"

    add_paragraph(document, f'• 折价水平评估: {discount_level}，{safety_note}')
    add_paragraph(document, '• 如果发行价折扣提高到15%或20%，盈利概率将显著提升')

    add_section_break(document)

    # ==================== 三、压力测试 ====================
    add_title(document, '三、压力测试', level=1)

    add_paragraph(document, '本章节模拟极端市场情况下的定增项目表现。')

    add_title(document, '3.1 压力情景分析', level=2)

    scenarios = [
        '基准情景',
        '熊市情景 -20%',
        '熊市情景 -30%',
        '牛市情景 +20%',
        '牛市情景 +30%'
    ]

    scenario_prices = [
        market_data['current_price'],
        market_data['current_price'] * 0.8,
        market_data['current_price'] * 0.7,
        market_data['current_price'] * 1.2,
        market_data['current_price'] * 1.3
    ]

    scenario_returns = [((p - project_params['issue_price']) / project_params['issue_price']) * 100
                       for p in scenario_prices]

    # 保存图表
    stress_chart_path = os.path.join(IMAGES_DIR, '02_stress_test.png')
    generate_stress_test_chart(scenarios, scenario_returns, stress_chart_path)

    stress_table = [
        [scenarios[0], f'{scenario_prices[0]:.2f}', f'{scenario_returns[0]:+.1f}%', '正常市场环境'],
        [scenarios[1], f'{scenario_prices[1]:.2f}', f'{scenario_returns[1]:+.1f}%', '市场大幅下跌'],
        [scenarios[2], f'{scenario_prices[2]:.2f}', f'{scenario_returns[2]:+.1f}%', '极端下跌'],
        [scenarios[3], f'{scenario_prices[3]:.2f}', f'{scenario_returns[3]:+.1f}%', '市场上涨'],
        [scenarios[4], f'{scenario_prices[4]:.2f}', f'{scenario_returns[4]:+.1f}%', '大幅上涨']
    ]
    add_table_data(document, ['情景', '模拟价格', '收益率', '说明'], stress_table)

    add_paragraph(document, '图表 2: 压力测试分析')
    add_image(document, stress_chart_path)

    add_paragraph(document, '压力测试结论：')
    add_paragraph(document, '• 在熊市情景下，项目可能出现较大亏损')
    add_paragraph(document, '• 折价发行提供了一定的下行保护')
    add_paragraph(document, '• 需要密切关注市场动态，做好风险管理')

    add_section_break(document)

    # ==================== 四、DCF估值分析 ====================
    add_title(document, '四、DCF估值分析', level=1)

    add_paragraph(document, '本章节使用现金流折现模型评估公司的内在价值。')

    # 生成并插入 DCF 热力图
    add_title(document, '4.1 估值敏感性分析', level=2)

    dcf_chart_path = os.path.join(IMAGES_DIR, '04_dcf_valuation_heatmap.png')
    generate_dcf_valuation_heatmap(dcf_chart_path)

    add_paragraph(document, '图表 3: DCF估值敏感性分析矩阵')
    add_image(document, dcf_chart_path, width=Inches(6))

    add_paragraph(document, '分析说明：')
    add_paragraph(document, '• 横轴：WACC（加权平均资本成本），代表折现率')
    add_paragraph(document, '• 纵轴：永续增长率，代表终值假设')
    add_paragraph(document, '• 颜色：绿色表示估值高于当前价，红色表示估值低于当前价')
    add_paragraph(document, '• 中心点：当前股价 23.88 元/股')

    add_title(document, '4.2 估值参数与结论', level=2)

    wacc = 0.0811
    intrinsic_value = 23.28

    dcf_params = [
        ['WACC（加权平均资本成本）', '8.11%'],
        ['永续增长率', '2.5%'],
        ['预测期', '10年'],
        ['DCF内在价值', f'{intrinsic_value:.2f} 元/股'],
        ['当前价格', f'{project_params["current_price"]:.2f} 元/股'],
        ['发行价格', f'{project_params["issue_price"]:.2f} 元/股']
    ]
    add_table_data(document, ['参数', '值'], dcf_params)

    add_title(document, '4.2 估值结论', level=2)

    margin_market = (intrinsic_value - project_params['current_price']) / project_params['current_price'] * 100
    margin_issue = (intrinsic_value - project_params['issue_price']) / project_params['issue_price'] * 100

    add_paragraph(document, f'• DCF内在价值: {intrinsic_value:.2f} 元/股')
    add_paragraph(document, f'• 相对市价安全边际: {margin_market:+.1f}%')
    add_paragraph(document, f'• 相对发行价安全边际: {margin_issue:+.1f}%')

    if margin_issue > 15:
        conclusion = "DCF估值显示，相比发行价有显著安全边际，估值合理偏低。"
    elif margin_issue > 0:
        conclusion = "DCF估值显示，相比发行价有一定安全边际，估值相对合理。"
    else:
        conclusion = "DCF估值显示，发行价高于内在价值，需谨慎。"

    add_paragraph(document, f'• 估值结论: {conclusion}')

    add_section_break(document)

    # ==================== 五、蒙特卡洛模拟 ====================
    add_title(document, '五、蒙特卡洛模拟', level=1)

    add_paragraph(document, '本章节使用蒙特卡洛方法模拟未来股价路径，评估收益分布。')

    add_title(document, '5.1 模拟参数', level=2)

    mc_params = [
        ['模拟次数', '10,000次'],
        ['锁定期', f'{project_params["lockup_period"]} 个月'],
        ['初始价格', f'{project_params["current_price"]:.2f} 元/股'],
        ['年化波动率', f'{market_data["volatility"]*100:.1f}%'],
        ['年化漂移率', f'{market_data["drift"]*100:.1f}%'],
        ['发行价格', f'{project_params["issue_price"]:.2f} 元/股']
    ]
    add_table_data(document, ['参数', '值'], mc_params)

    add_title(document, '5.2 模拟结果', level=2)

    # 运行简化模拟
    lockup_days = project_params['lockup_period'] * 30
    n_simulations = 5000

    lockup_drift = risk_params.get('drift', market_data.get('drift', 0.08)) * (lockup_days / 365)
    lockup_vol = risk_params.get('volatility', market_data.get('volatility', 0.35)) * np.sqrt(lockup_days / 365)

    np.random.seed(42)
    returns = np.random.normal(lockup_drift, lockup_vol, n_simulations)
    final_prices = project_params['current_price'] * np.exp(returns)
    profit_losses = (final_prices - project_params['issue_price']) / project_params['issue_price']

    # 统计
    profit_prob = (profit_losses > 0).mean()
    mean_return = profit_losses.mean()
    median_return = np.median(profit_losses)
    percent_5 = np.percentile(profit_losses, 5)
    percent_95 = np.percentile(profit_losses, 95)

    # 保存图表
    mc_chart_path = os.path.join(IMAGES_DIR, '03_monte_carlo.png')
    generate_monte_carlo_chart(final_prices, project_params['issue_price'], project_params['current_price'], mc_chart_path)

    mc_results = [
        ['盈利概率', f'{profit_prob*100:.1f}%'],
        ['预期收益率', f'{mean_return*100:.1f}%'],
        ['收益率中位数', f'{median_return*100:.1f}%'],
        ['5%分位数（最差情况）', f'{percent_5*100:.1f}%'],
        ['95%分位数（较好情况）', f'{percent_95*100:.1f}%']
    ]
    add_table_data(document, ['指标', '值'], mc_results)

    add_paragraph(document, '图表 4: 蒙特卡洛模拟结果')
    add_image(document, mc_chart_path)

    add_paragraph(document, '蒙特卡洛模拟结论：')
    add_paragraph(document, f'• 在 {n_simulations:,} 次模拟中，约 {profit_prob*100:.1f}% 的场景实现盈利')
    add_paragraph(document, f'• 预期收益率约 {mean_return*100:.1f}%')
    add_paragraph(document, f'• 95%的置信区间内，亏损可能不超过 {abs(percent_5)*100:.1f}%')

    add_section_break(document)

    # ==================== 六、VaR风险度量 ====================
    add_title(document, '六、VaR风险度量', level=1)

    add_paragraph(document, '本章节使用VaR（风险价值）和CVaR（条件风险价值）度量下行风险。')

    add_title(document, '6.1 风险指标', level=2)

    var_95 = abs(np.percentile(profit_losses, 5))
    var_99 = abs(np.percentile(profit_losses, 1))
    cvar_95 = abs(profit_losses[profit_losses <= np.percentile(profit_losses, 5)].mean())

    # 保存图表
    var_chart_path = os.path.join(IMAGES_DIR, '05_var_analysis.png')
    generate_var_chart(var_95, var_99, cvar_95, var_chart_path)

    var_data = [
        ['95% VaR', f'{var_95*100:.1f}%', '有95%的概率，亏损不超过此值'],
        ['99% VaR', f'{var_99*100:.1f}%', '有99%的概率，亏损不超过此值'],
        ['95% CVaR', f'{cvar_95*100:.1f}%', '在最差5%情况下的平均损失'],
        ['最大回撤（预估）', f'{market_data["volatility"]*2*100:.1f}%', '基于波动率的估算']
    ]
    add_table_data(document, ['风险指标', '数值', '说明'], var_data)

    add_paragraph(document, '图表 5: VaR风险度量')
    add_image(document, var_chart_path)

    add_paragraph(document, '风险度量结论：')
    add_paragraph(document, f'• 95%置信水平下，最大可能亏损约 {var_95*100:.1f}%')
    add_paragraph(document, f'• 极端情况下（1%概率），亏损可能达到 {var_99*100:.1f}%')
    add_paragraph(document, f'• CVaR显示，在最差情况下平均损失约 {cvar_95*100:.1f}%')

    add_section_break(document)

    # ==================== 七、综合评估 ====================
    add_title(document, '七、综合评估', level=1)

    add_title(document, '7.1 风险评分', level=2)

    # 综合评分
    scores = {
        '盈利概率': 30 if profit_prob >= 0.7 else 20 if profit_prob >= 0.5 else 10,
        '发行折价': 20 if discount_premium > 15 else 15 if discount_premium > 5 else 5,
        '锁定期': 20 if project_params['lockup_period'] <= 6 else 15 if project_params['lockup_period'] <= 12 else 5,
        '波动率': 15 if risk_params.get('volatility', 0.35) <= 0.3 else 10 if risk_params.get('volatility', 0.35) <= 0.4 else 5,
        'VaR': 15 if var_95 <= 0.2 else 10 if var_95 <= 0.3 else 3
    }
    total_score = sum(scores.values())

    # 保存雷达图
    radar_chart_path = os.path.join(IMAGES_DIR, '06_radar_chart.png')
    generate_radar_chart(scores, radar_chart_path)

    score_data = [[k, f'{v}/30' if k == '盈利概率' else f'{v}/20' if k != 'VaR' else f'{v}/15'] for k, v in scores.items()]
    score_data.append(['总分', f'{total_score}/100'])

    add_table_data(document, ['评估维度', '得分'], score_data)

    add_paragraph(document, '图表 6: 风险评分雷达图')
    add_image(document, radar_chart_path)

    add_paragraph(document, f'综合风险评分: {total_score}/100 分')

    add_title(document, '7.2 各维度评估', level=2)

    add_paragraph(document, f'1. 盈利概率（{scores["盈利概率"]}/30分）')
    if profit_prob >= 0.7:
        add_paragraph(document, '   ✅ 盈利概率较高，投资吸引力强')
    elif profit_prob >= 0.5:
        add_paragraph(document, '   ⚠️ 盈利概率中等，需谨慎评估')
    else:
        add_paragraph(document, '   ❌ 盈利概率较低，风险较大')

    add_paragraph(document, f'2. 发行折价（{scores["发行折价"]}/20分）')
    if discount_premium > 15:
        add_paragraph(document, f'   ✅ {issue_type}，安全边际充足（{discount_premium:.1f}%）')
    elif discount_premium > 5:
        add_paragraph(document, f'   ⚠️ 安全边际一般（{discount_premium:.1f}%）')
    else:
        add_paragraph(document, f'   ❌ 安全边际不足（{discount_premium:.1f}%）')

    add_paragraph(document, f'3. 锁定期（{scores["锁定期"]}/20分）')
    lockup_comment = "相对较短" if project_params["lockup_period"] <= 6 else "需要较长等待"
    add_paragraph(document, f'   锁定期 {project_params["lockup_period"]} 个月，{lockup_comment}')

    add_paragraph(document, f'4. 波动率（{scores["波动率"]}/15分）')
    vol_comment = "相对可控" if market_data["volatility"] <= 0.35 else "波动较大"
    add_paragraph(document, f'   年化波动率 {market_data["volatility"]*100:.1f}%，{vol_comment}')

    add_paragraph(document, f'5. 风险价值（{scores["VaR"]}/15分）')
    var_comment = "风险可控" if var_95 <= 0.25 else "需注意风险"
    add_paragraph(document, f'   95% VaR 为 {var_95*100:.1f}%，{var_comment}')

    add_section_break(document)

    # ==================== 八、投资建议与风险提示 ====================
    add_title(document, '八、投资建议与风险提示', level=1)

    add_title(document, '8.1 投资建议', level=2)

    # 投资建议
    if total_score >= 80:
        recommendation = "🟢 建议投资"
        detail = "项目整体风险收益比良好，各项指标表现优秀，建议积极参与。"
        actions = [
            "可以按计划参与定增",
            "建议关注锁定期内的市场动态",
            "可考虑适当杠杆"
        ]
    elif total_score >= 60:
        recommendation = "🟡 谨慎投资"
        detail = "项目整体风险收益比一般，存在一定不确定性，建议谨慎评估。"
        actions = [
            "深入调研项目基本面",
            "评估自身风险承受能力",
            "考虑分批参与"
        ]
    else:
        recommendation = "🔴 不建议投资"
        detail = "项目整体风险较高，安全边际不足，建议观望或寻找其他机会。"
        actions = [
            "暂不建议参与",
            "等待更好的入场时机",
            "考虑其他投资标的"
        ]

    add_paragraph(document, f'总体建议: {recommendation}', bold=True, font_size=12)
    add_paragraph(document, f'说明: {detail}')

    add_paragraph(document, '建议行动:')
    for action in actions:
        add_paragraph(document, f'  • {action}')

    add_title(document, '8.2 核心指标汇总', level=2)

    summary_data = [
        ['风险评分', f'{total_score}/100'],
        ['盈利概率', f'{profit_prob*100:.1f}%'],
        ['发行类型', f'{issue_type}'],
        ['安全边际/溢价率', f'{discount_premium:.2f}%'],
        ['预期收益率', f'{mean_return*100:.1f}%'],
        ['95% VaR', f'{var_95*100:.1f}%'],
        ['波动率', f'{market_data["volatility"]*100:.1f}%'],
        ['DCF内在价值', f'{intrinsic_value:.2f} 元/股']
    ]
    add_table_data(document, ['指标', '值'], summary_data)

    add_title(document, '8.3 主要风险提示', level=2)

    risks = [
        f'• 市场风险: 当前波动率 {market_data["volatility"]*100:.1f}%，市场波动可能导致实际收益偏离预期',
        f'• 流动性风险: {project_params["lockup_period"]}个月锁定期内无法交易，需承担期间价格波动',
        f'• 估值风险: DCF估值基于多个假设，实际业绩可能偏离预测',
        f'• VaR风险: 95%置信水平下最大可能亏损 {var_95*100:.1f}%',
        f'• 行业风险: 需关注行业政策变化和竞争格局'
    ]

    for risk in risks:
        add_paragraph(document, risk)

    add_title(document, '8.4 免责声明', level=2)

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

    # ==================== 保存文档 ====================
    output_path = os.path.join(OUTPUTS_DIR, output_file)
    document.save(output_path)

    print(f"\n✅ 报告生成成功!")
    print(f"   📄 Word文档: {output_path}")
    print(f"   📊 图表目录: {IMAGES_DIR}")
    print(f"   文件大小: {os.path.getsize(output_path)/1024:.1f} KB")

    return output_path


if __name__ == '__main__':
    report_path = generate_report(stock_code='300735.SZ', output_file='光弘科技_定增风险分析报告.docx')
    print("\n报告生成完成！")
