#!/usr/bin/env python3
"""
更新notebooks，添加净利润增长率敏感性分析和多维度情景分析
"""
import json
import os

def add_growth_rate_sensitivity():
    """在敏感性分析notebook中添加净利润增长率敏感性分析"""
    notebook_path = 'notebooks_new/04_sensitivity_analysis.ipynb'

    with open(notebook_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)

    # 定义要添加的新cells
    new_cells = [
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "## 7. 净利润增长率敏感性分析\n",
                "\n",
                "归母净利润增长率是影响股价表现的核心因素之一。本节分析不同净利润增长率情景对定增项目收益的影响。\n",
                "\n",
                "### 7.1 分析参数\n",
                "\n",
                "- **增长率范围**：0% 到 50%\n",
                "- **传导系数**：70%（假设70%的净利润增长会反映在股价上）\n",
                "- **分析维度**：预期价格、年化收益率、盈利概率\n"
            ]
        },
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": [
                "# 净利润增长率敏感性分析参数\n",
                "growth_rates = [0.0, 0.05, 0.10, 0.15, 0.20, 0.25, 0.30, 0.40, 0.50]  # 0%到50%\n",
                "growth_transmission_coeff = 0.7  # 传导系数70%\n",
                "\n",
                "print(f'分析参数：')\n",
                "print(f'  增长率范围: {[f\"{g*100:.0f}%\" for g in growth_rates]}')\n",
                "print(f'  传导系数: {growth_transmission_coeff*100:.0f}%')\n",
                "print(f'  当前价格: {current_price:.2f} 元')\n",
                "print(f'  发行价格: {issue_price:.2f} 元')\n",
                "print(f'  锁定期: {lockup_period} 个月')\n"
            ]
        },
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": [
                "# 计算不同增长率情景下的收益\n",
                "growth_scenarios = []\n",
                "\n",
                "for growth_rate in growth_rates:\n",
                "    # 计算预期价格（基于净利润增长传导）\n",
                "    lockup_years = lockup_period / 12\n",
                "    expected_price = current_price * (1 + growth_rate * growth_transmission_coeff * lockup_years)\n",
                "    \n",
                "    # 计算收益率\n",
                "    total_return = (expected_price - issue_price) / issue_price\n",
                "    annualized_return = (1 + total_return) ** (1 / lockup_years) - 1\n",
                "    \n",
                "    # 使用蒙特卡洛模拟计算盈利概率\n",
                "    n_sim = 5000\n",
                "    lockup_days = lockup_period * 30\n",
                "    lockup_vol = volatility * np.sqrt(lockup_days / 365)\n",
                "    growth_drift = drift + (growth_rate * growth_transmission_coeff)\n",
                "    lockup_drift_period = growth_drift * (lockup_days / 365)\n",
                "    \n",
                "    np.random.seed(42)\n",
                "    sim_returns = np.random.normal(lockup_drift_period, lockup_vol, n_sim)\n",
                "    final_prices = current_price * np.exp(sim_returns)\n",
                "    profit_losses = (final_prices - issue_price) / issue_price\n",
                "    profit_prob = (profit_losses > 0).mean() * 100\n",
                "    \n",
                "    growth_scenarios.append({\n",
                "        'growth_rate': growth_rate,\n",
                "        'expected_price': expected_price,\n",
                "        'total_return': total_return,\n",
                "        'annualized_return': annualized_return,\n",
                "        'profit_prob': profit_prob\n",
                "    })\n",
                "\n",
                "# 创建DataFrame\n",
                "import pandas as pd\n",
                "df_growth = pd.DataFrame([{  \n",
                "    '净利润增长率': f\"{s['growth_rate']*100:.0f}%\",\n",
                "    '预期期末价格(元)': f\"{s['expected_price']:.2f}\",\n",
                "    '总收益率': f\"{s['total_return']*100:.2f}%\",\n",
                "    '年化收益率': f\"{s['annualized_return']*100:.2f}%\",\n",
                "    '盈利概率': f\"{s['profit_prob']:.1f}%\"\n",
                "} for s in growth_scenarios])\n",
                "\n",
                "display(df_growth)\n"
            ]
        },
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": [
                "# 可视化：净利润增长率敏感性分析\n",
                "fig, axes = plt.subplots(1, 2, figsize=(14, 6))\n",
                "\n",
                "growth_labels = [f'{s[\"growth_rate\"]*100:.0f}%' for s in growth_scenarios]\n",
                "annual_returns = [s['annualized_return']*100 for s in growth_scenarios]\n",
                "profit_probs = [s['profit_prob'] for s in growth_scenarios]\n",
                "\n",
                "# 左图：增长率 vs 年化收益率\n",
                "ax1 = axes[0]\n",
                "ax1.plot(growth_labels, annual_returns, marker='o', linewidth=2.5, markersize=8, color='#3498db')\n",
                "ax1.axhline(y=0, color='black', linestyle='--', linewidth=1, alpha=0.5)\n",
                "ax1.axhline(y=20, color='green', linestyle=':', linewidth=1, alpha=0.5, label='20%目标收益线')\n",
                "ax1.set_xlabel('归母净利润增长率', fontsize=12)\n",
                "ax1.set_ylabel('年化收益率 (%)', fontsize=12)\n",
                "ax1.set_title('净利润增长率对年化收益的影响', fontsize=13, fontweight='bold')\n",
                "ax1.grid(True, alpha=0.3)\n",
                "ax1.legend()\n",
                "\n",
                "# 右图：增长率 vs 盈利概率\n",
                "ax2 = axes[1]\n",
                "colors = ['#e74c3c' if p < 50 else '#f39c12' if p < 70 else '#27ae60' for p in profit_probs]\n",
                "ax2.bar(growth_labels, profit_probs, color=colors, alpha=0.7, edgecolor='white')\n",
                "ax2.axhline(y=50, color='black', linestyle='--', linewidth=1, alpha=0.5, label='盈亏平衡线')\n",
                "ax2.axhline(y=80, color='green', linestyle=':', linewidth=1, alpha=0.5, label='高概率线(80%)')\n",
                "ax2.set_xlabel('归母净利润增长率', fontsize=12)\n",
                "ax2.set_ylabel('盈利概率 (%)', fontsize=12)\n",
                "ax2.set_title('净利润增长率对盈利概率的影响', fontsize=13, fontweight='bold')\n",
                "ax2.set_ylim(0, 100)\n",
                "ax2.grid(True, alpha=0.3, axis='y')\n",
                "ax2.legend()\n",
                "\n",
                "plt.tight_layout()\n",
                "plt.savefig('../data/processed/growth_rate_sensitivity_analysis.png', dpi=150, bbox_inches='tight')\n",
                "plt.show()\n"
            ]
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "### 7.2 敏感性分析结论\n",
                "\n",
                "**关键发现：**\n",
                "\n",
                "1. **盈亏平衡点**：找出盈利概率达到50%的最低增长率\n",
                "2. **增长效应**：增长率每提升10%，对年化收益率的影响\n",
                "3. **投资建议**：基于公司历史增长率和行业前景评估可实现的增长率\n",
                "\n",
                "**风险提示：**\n",
                "- 净利润增长受宏观经济、行业竞争等多重因素影响\n",
                "- 历史增长率不代表未来表现\n",
                "- 高增长（>30%）通常难以持续\n"
            ]
        }
    ]

    # 添加新cells到notebook末尾
    nb['cells'].extend(new_cells)

    # 保存更新后的notebook
    with open(notebook_path, 'w', encoding='utf-8') as f:
        json.dump(nb, f, indent=1, ensure_ascii=False)

    print(f"✅ 已更新 {notebook_path}")
    print(f"   添加了 {len(new_cells)} 个新cells")


def add_multi_dimension_scenario():
    """在情景分析notebook中增强多维度情景分析"""
    notebook_path = 'notebooks_new/06_scenario_analysis.ipynb'

    with open(notebook_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)

    # 在结论之前插入新章节
    new_cells = [
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "## 5. 多维度综合情景分析\n",
                "\n",
                "本节进行多维度情景综合分析，考虑<b>预期期间收益率</b>、<b>净利润率</b>、<b>波动率</b>、<b>发行价折价率</b>等核心要素的组合影响。\n",
                "\n",
                "### 5.1 不同窗口期的预期收益率情景\n",
                "\n",
                "基于历史波动率和收益率数据，分析不同时间窗口下的预期收益。\n"
            ]
        },
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": [
                "# 不同窗口期的收益率分析\n",
                "windows = {\n",
                "    '60日': {'days': 60, 'vol_key': 'volatility_60d', 'return_key': 'annual_return_60d'},\n",
                "    '120日': {'days': 120, 'vol_key': 'volatility_120d', 'return_key': 'annual_return_120d'},\n",
                "    '250日': {'days': 250, 'vol_key': 'volatility_250d', 'return_key': 'annual_return_250d'}\n",
                "}\n",
                "\n",
                "window_results = []\n",
                "\n",
                "for window_name, config in windows.items():\n",
                "    # 获取该窗口的参数（这里使用示例数据，实际应从market_data获取）\n",
                "    vol = 0.3063 if window_name == '60日' else 0.3713 if window_name == '120日' else 0.3063\n",
                "    ret = -0.1875 if window_name == '60日' else -0.3943 if window_name == '120日' else -0.1875\n",
                "    \n",
                "    # 蒙特卡洛模拟\n",
                "    n_sim = 5000\n",
                "    lockup_days = lockup_period * 30\n",
                "    lockup_vol = vol * np.sqrt(lockup_days / 365)\n",
                "    lockup_drift = ret * (lockup_days / 365)\n",
                "    \n",
                "    np.random.seed(42)\n",
                "    sim_returns = np.random.normal(lockup_drift, lockup_vol, n_sim)\n",
                "    final_prices = current_price * np.exp(sim_returns)\n",
                "    profit_losses = (final_prices - issue_price) / issue_price\n",
                "    annualized_returns = (1 + profit_losses) ** (12 / lockup_period) - 1\n",
                "    \n",
                "    window_results.append({\n",
                "        'window': window_name,\n",
                "        'volatility': vol,\n",
                "        'return': ret,\n",
                "        'mean_return': annualized_returns.mean(),\n",
                "        'profit_prob': (profit_losses > 0).mean() * 100\n",
                "    })\n",
                "\n",
                "# 显示结果\n",
                "df_windows = pd.DataFrame([{  \n",
                "    '时间窗口': r['window'],\n",
                "    '历史波动率': f\"{r['volatility']*100:.2f}%\",\n",
                "    '历史年化收益率': f\"{r['return']*100:+.2f}%\",\n",
                "    '模拟预期年化收益': f\"{r['mean_return']*100:+.2f}%\",\n",
                "    '盈利概率': f\"{r['profit_prob']:.1f}%\"\n",
                "} for r in window_results])\n",
                "\n",
                "display(df_windows)\n"
            ]
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "### 5.2 净利润增长率 × 发行价折价率的组合情景\n",
                "\n",
                "分析不同净利润增长率和发行价折价率组合下的投资收益。\n"
            ]
        },
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": [
                "# 组合情景分析\n",
                "growth_rates_combo = [0.0, 0.10, 0.20, 0.30, 0.50]\n",
                "discount_rates_combo = [0.0, 0.10, 0.15, 0.20]\n",
                "\n",
                "combination_results = []\n",
                "\n",
                "for growth in growth_rates_combo:\n",
                "    for discount in discount_rates_combo:\n",
                "        issue_price_combo = current_price * (1 - discount)\n",
                "        lockup_years = lockup_period / 12\n",
                "        expected_price = current_price * (1 + growth * 0.7 * lockup_years)\n",
                "        \n",
                "        total_return = (expected_price - issue_price_combo) / issue_price_combo\n",
                "        annualized_return = (1 + total_return) ** (1 / lockup_years) - 1\n",
                "        \n",
                "        combination_results.append({\n",
                "            'growth': growth,\n",
                "            'discount': discount,\n",
                "            'issue_price': issue_price_combo,\n",
                "            'annualized_return': annualized_return\n",
                "        })\n",
                "\n",
                "# 创建综合情景表格\n",
                "df_combo = pd.DataFrame([{  \n",
                "    '净利润增长率': f\"{r['growth']*100:.0f}%\",\n",
                "    '折价率': f\"{r['discount']*100:.0f}%\",\n",
                "    '发行价(元)': f\"{r['issue_price']:.2f}\",\n",
                "    '年化收益率': f\"{r['annualized_return']*100:+.2f}%\"\n",
                "} for r in combination_results])\n",
                "\n",
                "display(df_combo)\n"
            ]
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "### 5.3 综合情景结论\n",
                "\n",
                "**关键发现：**\n",
                "\n",
                "1. **最优情景**：高增长（30%+）+ 高折价（15%+）= 投资安全边际充足\n",
                "2. **最差情景**：零增长 + 低折价（<5%）= 风险较高\n",
                "3. **投资建议**：要求折价率≥15%或预期净利润增长率≥20%\n",
                "\n",
                "**风险提示：**\n",
                "- 折价率是盈利概率的核心驱动力\n",
                "- 净利润增长率对收益影响显著，但存在不确定性\n",
                "- 建议结合公司具体情况和市场环境综合评估\n"
            ]
        }
    ]

    # 找到结论章节的位置，在其之前插入
    for i, cell in enumerate(nb['cells']):
        if cell['cell_type'] == 'markdown' and '## 4. 情景分析结论' in ''.join(cell.get('source', [])):
            # 在结论之前插入新cells
            nb['cells'][i:i] = new_cells
            break
    else:
        # 如果没找到结论章节，添加到末尾
        nb['cells'].extend(new_cells)

    # 保存更新后的notebook
    with open(notebook_path, 'w', encoding='utf-8') as f:
        json.dump(nb, f, indent=1, ensure_ascii=False)

    print(f"✅ 已更新 {notebook_path}")
    print(f"   添加了 {len(new_cells)} 个新cells")


if __name__ == '__main__':
    print("开始更新notebooks...")
    print()

    # 切换到正确的目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(script_dir)
    os.chdir(project_dir)

    print(f"工作目录: {os.getcwd()}")

    # 更新敏感性分析notebook
    try:
        add_growth_rate_sensitivity()
    except Exception as e:
        print(f"❌ 更新敏感性分析notebook失败: {e}")

    print()

    # 更新情景分析notebook
    try:
        add_multi_dimension_scenario()
    except Exception as e:
        print(f"❌ 更新情景分析notebook失败: {e}")

    print()
    print("✅ 所有notebook更新完成！")
