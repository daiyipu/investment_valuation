"""
附件：585种情景完整数据表

本附件提供完整的585种情景组合数据，供备查参考。
情景组合：漂移率（-30%~+30%，13档）× 波动率（10%~50%，5档）× 溢价率（-20%~+20%，9档）= 585种
"""

def generate_chapter(context):
    """
    生成附件：585种情景完整数据表（统一接口）

    参数:
        context: 包含document和results的字典
    """
    document = context['document']
    all_scenarios_for_appendix = context['results'].get('all_scenarios', [])
    _generate_appendix_scenarios(document, all_scenarios_for_appendix)
    return context


def _generate_appendix_scenarios(document, all_scenarios_for_appendix):
    """
    生成附件：585种情景完整数据表

    参数说明：
    - document: Word文档对象
    - all_scenarios_for_appendix: 所有情景数据列表，每个元素为字典包含：
        - drift: 漂移率
        - volatility: 波动率
        - discount: 溢价率
        - issue_price: 发行价
        - mean_return: 预期年化收益率
        - median_return: 中位数收益率
        - profit_prob: 盈利概率
        - var_5: 5% VaR
        - var_95: 95% VaR
    """
    from module_utils import add_title, add_paragraph, add_table_data, add_section_break

    # ==================== 附件：585种情景完整数据表 ====================
    add_section_break(document)
    add_title(document, '附件：585种情景完整数据表', level=1)

    add_paragraph(document, '本附件提供完整的585种情景组合数据，供备查参考。')
    add_paragraph(document, f'情景组合：漂移率（-30%~+30%，13档）× 波动率（10%~50%，5档）× 溢价率（-20%~+20%，9档）= {len(all_scenarios_for_appendix)}种')
    add_paragraph(document, '')
    add_paragraph(document, '说明：本表按波动率分块展示，便于根据实际市场波动率选择参考情景。在每个波动率区间内，按漂移率从高到低排序（漂移率相同时按溢价率从高到低排序）。建议参考创业板180日波动率（约36%）确定目标波动率区间。')
    add_paragraph(document, '')

    # 按波动率分块展示
    # 定义波动率区间
    vol_ranges = [
        ('低波动率区间 (10%-20%)', 10, 20),
        ('中低波动率区间 (20%-30%)', 20, 30),
        ('中高波动率区间 (30%-40%)', 30, 40),
        ('高波动率区间 (40%-50%)', 40, 50)
    ]

    for vol_name, vol_min, vol_max in vol_ranges:
        # 筛选当前波动率区间的情景
        # 兼容新旧格式：新格式将数据嵌套在'scenario'键中
        scenarios_in_range = []
        for s in all_scenarios_for_appendix:
            # 获取波动率值（兼容嵌套和扁平结构）
            if 'volatility' in s:
                vol = s['volatility']
            elif 'scenario' in s and 'volatility' in s['scenario']:
                vol = s['scenario']['volatility']
            else:
                continue

            if vol_min <= vol*100 < vol_max:
                scenarios_in_range.append(s)

        if not scenarios_in_range:
            continue

        # 在区间内按漂移率倒序排列，漂移率相同时按溢价率倒序排列
        # 漂移率倒序：从高到低（+30%到-30%）
        # 溢价率倒序：从高到低（+20%到-20%），即从溢价到折价

        def get_sort_key(x):
            # 兼容嵌套和扁平结构
            if 'drift' in x:
                drift = x['drift']
                discount = x.get('discount', x.get('premium_rate', 0))
            elif 'scenario' in x:
                drift = x['scenario']['drift']
                discount = x['scenario'].get('discount', x['scenario'].get('premium_rate', 0))
            else:
                drift = 0
                discount = 0
            return (-drift, -discount)

        scenarios_sorted = sorted(scenarios_in_range, key=get_sort_key)

        # 添加区块标题
        add_title(document, f'附表：{vol_name}', level=2)

        # 生成表格数据
        appendix_data = []
        for i, s in enumerate(scenarios_sorted, 1):
            # 兼容嵌套和扁平结构
            if 'drift' in s:
                # 扁平结构
                drift = s['drift']
                volatility = s['volatility']
                discount = s.get('discount', s.get('premium_rate', 0))
                issue_price = s.get('issue_price', 0)
                mean_return = s.get('mean_return', 0)
                median_return = s.get('median_return', 0)
                profit_prob = s.get('profit_prob', 0)
                var_5 = s.get('var_5', 0)
                var_95 = s.get('var_95', 0)
            else:
                # 嵌套结构
                scenario = s['scenario']
                drift = scenario['drift']
                volatility = scenario['volatility']
                discount = scenario.get('discount', scenario.get('premium_rate', 0))
                issue_price = scenario.get('issue_price', 0)
                mean_return = s.get('mean_return', 0)
                median_return = s.get('median_return', 0)
                profit_prob = s.get('profit_prob', 0)
                var_5 = s.get('var_5', 0)
                var_95 = s.get('var_95', 0)

            appendix_data.append([
                f"{i}",
                f"{drift*100:+.0f}%",
                f"{volatility*100:.0f}%",
                f"{discount*100:+.0f}%",
                f"{issue_price:.2f}",
                f"{mean_return*100:+.2f}%",
                f"{median_return*100:+.2f}%",
                f"{profit_prob:.1f}%",
                f"{var_5*100:+.2f}%",
                f"{var_95*100:+.2f}%"
            ])

        appendix_headers = ['排名', '漂移率', '波动率', '溢价率', '发行价(元)', '预期年化收益', '中位数收益', '盈利概率', '5% VaR', '95% VaR']
        add_table_data(document, appendix_headers, appendix_data)
        add_paragraph(document, '')

    add_paragraph(document, '')
    add_paragraph(document, '附表说明：')
    add_paragraph(document, '• 排序方式：')
    add_paragraph(document, '  - 按波动率分块（低波动、中低波动、中高波动、高波动）')
    add_paragraph(document, '  - 每个区块内按漂移率从高到低排序（+30%到-30%）')
    add_paragraph(document, '  - 漂移率相同时按溢价率从高到低排序（+20%到-20%）')
    add_paragraph(document, '• 漂移率：年化收益率，反映股价预期趋势（正值=上涨趋势，负值=下跌趋势）')
    add_paragraph(document, '• 波动率：年化波动率，反映股价不确定性')
    add_paragraph(document, '• 溢价率：发行价相对MA20的溢价（负值=折价，正值=溢价）')
    add_paragraph(document, '• 预期年化收益：基于蒙特卡洛模拟的平均年化收益率')
    add_paragraph(document, '• 中位数收益：年化收益率的中位数')
    add_paragraph(document, '• 盈利概率：到期时盈利（收益率>0）的概率')
    add_paragraph(document, '• 5% VaR：95%的情景下收益率不低于此值（较好情况）')
    add_paragraph(document, '• 95% VaR：5%的情景下收益率不高于此值（最差情况）')
    add_paragraph(document, '')
    add_paragraph(document, '使用建议：')
    add_paragraph(document, f'• 参考创业板180日波动率（约36%），优先查看"中高波动率区间 (30%-40%)"的情景')
    add_paragraph(document, '• 根据对未来市场走势的判断（乐观/中性/悲观），在相应漂移率区间查找情景')
    add_paragraph(document, '• 结合当前项目的实际溢价率，找到最接近的情景作为参考基准')
    add_paragraph(document, '• 对比不同情景下的预期收益和盈利概率，评估投资价值')
