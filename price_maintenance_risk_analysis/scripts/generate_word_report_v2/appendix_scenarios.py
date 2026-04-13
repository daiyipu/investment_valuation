"""
附件：情景分析完整数据表

本附件提供所有情景组合数据，供备查参考。
包含：
- 6.1 多参数情景：585种（13×5×9）
- 6.2 市场指数情景：45种（3×3×5）
- 6.3 行业PE情景：45种（3×3×5）
- 6.4 个股PE情景：45种（3×3×5）
- 6.5 DCF估值情景：15种（1×3×5）
"""

def generate_chapter(context):
    """
    生成附件：情景分析完整数据表（统一接口）

    参数:
        context: 包含document和results的字典
    """
    document = context['document']
    # 分别获取两种情景数据
    multi_param_scenarios = context['results'].get('multi_param_scenarios_585', [])
    historical_scenarios = context['results'].get('historical_scenarios_195', [])
    project_params = context['project_params']
    market_data = context['market_data']

    # 生成附件内容
    _generate_appendix_scenarios(document, multi_param_scenarios, historical_scenarios, project_params, market_data)
    return context


def _generate_appendix_scenarios(document, multi_param_scenarios, historical_scenarios, project_params, market_data):
    """
    生成附件：情景分析完整数据表

    参数说明：
    - document: Word文档对象
    - multi_param_scenarios: 6.1节的585种多参数构造情景
    - historical_scenarios: 6.2-6.5节的195种历史数据情景
    - project_params: 项目参数
    - market_data: 市场数据
    """
    from module_utils import add_title, add_paragraph, add_table_data, add_section_break

    # ==================== 附件1：585种多参数构造情景完整数据表 ====================
    add_section_break(document)
    add_title(document, '附件1：585种多参数构造情景完整数据表', level=1)

    add_paragraph(document, '本附件提供完整的585种多参数构造情景组合数据，供备查参考。')
    add_paragraph(document, f'情景组合：漂移率（-30%~+30%，13档）× 波动率（10%~50%，5档）× 溢价率（-20%~+20%，9档）= {len(multi_param_scenarios)}种')
    add_paragraph(document, '')
    add_paragraph(document, '说明：本表按波动率分块展示，便于根据实际市场波动率选择参考情景。在每个波动率区间内，按漂移率从高到低排序（漂移率相同时按溢价率从高到低排序）。建议参考创业板180日波动率（约36%）确定目标波动率区间。')
    add_paragraph(document, '')

    _generate_multi_param_tables(document, multi_param_scenarios)

    # ==================== 附件2：195种历史数据情景完整数据表 ====================
    add_section_break(document)
    add_title(document, '附件2：195种历史数据情景完整数据表', level=1)

    add_paragraph(document, '本附件提供6.2-6.5节的历史数据情景，包括市场指数情景、行业PE情景、个股PE情景和DCF估值情景。')
    add_paragraph(document, f'情景总数：{len(historical_scenarios)}种')
    add_paragraph(document, '')
    add_paragraph(document, '说明：历史数据情景基于实际历史数据构建，参数范围可能与多参数构造情景不同。')
    add_paragraph(document, '')

    _generate_historical_tables(document, historical_scenarios)

    # ==================== 附件3：溢价率对比表 ====================
    add_section_break(document)
    add_title(document, '附件3：溢价率对比表', level=1)

    add_paragraph(document, '本附件提供不同溢价率下的发行价对比，以及与发行日价格和当前价格的差异情况。')
    add_paragraph(document, '')

    _generate_premium_rate_table(document, project_params, market_data)

def _generate_multi_param_tables(document, multi_param_scenarios):
    """生成585种多参数构造情景的表格"""
    from module_utils import add_title, add_paragraph, add_table_data

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
        for s in multi_param_scenarios:
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
    add_paragraph(document, '')


def _generate_historical_tables(document, historical_scenarios):
    """生成195种历史数据情景的表格"""
    from module_utils import add_title, add_paragraph, add_table_data

    # ==================== 6.2-6.5节情景数据表 ====================
    add_title(document, '附表：6.2-6.5节情景数据表', level=2)
    add_paragraph(document, '本附表展示6.2至6.5节的专项情景分析结果，包括市场指数、行业PE、个股PE和DCF估值的情景数据。')
    add_paragraph(document, '')

    # 按情景类型分组展示
    scenario_groups = {
        '市场指数': '6.2',
        '行业PE': '6.3',
        '个股PE': '6.4',
        'DCF估值': '6.5'
    }

    # 按情景类型分组
    grouped_scenarios = {}
    for s in historical_scenarios:
        # 获取情景名称（兼容嵌套和扁平结构）
        scenario_name = None
        if 'name' in s:
            # 扁平结构
            scenario_name = s['name']
        elif 'scenario' in s and 'name' in s['scenario']:
            # 嵌套结构
            scenario_name = s['scenario']['name']

        # 如果没有name字段，跳过
        if not scenario_name:
            continue

        # 确定情景类型
        scenario_type = None
        for type_name in scenario_groups.keys():
            if type_name in scenario_name:
                scenario_type = type_name
                break

        if scenario_type:
            if scenario_type not in grouped_scenarios:
                grouped_scenarios[scenario_type] = []
            grouped_scenarios[scenario_type].append(s)

    # 检查是否有分组数据
    if not grouped_scenarios:
        add_paragraph(document, '未找到6.2-6.5节的专项情景数据，请确认情景分析章节已正确生成。')
        add_paragraph(document, '')
        return

    # 为每种情景类型生成表格
    for scenario_type, scenarios in grouped_scenarios.items():
        if not scenarios:
            continue

        # 添加类型标题
        section_num = scenario_groups[scenario_type]
        add_title(document, f'附表{section_num}：{scenario_type}情景数据表', level=3)

        # 按盈利概率排序
        scenarios_sorted = sorted(scenarios, key=lambda x: x.get('profit_prob', 0), reverse=True)

        # 生成表格数据（显示全部数据）
        table_data = []
        for i, s in enumerate(scenarios_sorted, 1):  # 显示全部数据
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
                scenario_name = s.get('name', '')
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
                scenario_name = scenario.get('name', '')

            table_data.append([
                f"{i}",
                scenario_name,
                f"{drift*100:+.1f}%",
                f"{volatility*100:.1f}%",
                f"{discount*100:+.0f}%",
                f"{issue_price:.2f}",
                f"{median_return*100:+.1f}%",
                f"{profit_prob:.1f}%"
            ])

        headers = ['排名', '情景名称', '漂移率', '波动率', '溢价率', '发行价(元)', '中位数收益', '盈利概率']
        add_table_data(document, headers, table_data)

        add_paragraph(document, f'注：{scenario_type}共{len(scenarios_sorted)}个情景，按盈利概率从高到低排序。')

        add_paragraph(document, '')
    add_paragraph(document, '附表说明（6.2-6.5）：')
    add_paragraph(document, '• 本表展示6.2至6.5节的专项情景分析结果')
    add_paragraph(document, '• 每种情景类型按盈利概率从高到低排序')
    add_paragraph(document, '• 市场指数情景：基于沪深300、中证500、创业板指、科创50的历史数据')
    add_paragraph(document, '• 行业PE情景：基于行业PE分位数的估值回归情景')
    add_paragraph(document, '• 个股PE情景：基于个股PE分位数的估值回归情景')
    add_paragraph(document, '• DCF估值情景：基于DCF内在价值的估值情景')


def _generate_premium_rate_table(document, project_params, market_data):
    """
    生成溢价率对比表

    参数说明：
    - document: Word文档对象
    - project_params: 项目参数
    - market_data: 市场数据
    """
    from module_utils import add_title, add_paragraph, add_table_data

    # 获取基础价格数据
    ma20_price = market_data.get('ma_20', 0)
    current_price = market_data.get('current_price', 0)
    issue_date_price = project_params.get('current_price', 0)  # 发行日价格
    actual_issue_price = project_params.get('issue_price', 0)  # 实际发行价

    # 获取发行日期
    issue_date = project_params.get('issue_date', 'N/A')

    # 生成从-20%到+10%的溢价率表格，每1%一个档
    add_title(document, '附表：溢价率与发行价对比表', level=2)

    add_paragraph(document, f'基础价格信息：')
    add_paragraph(document, f'• MA20价格：{ma20_price:.2f} 元/股')
    add_paragraph(document, f'• 当前价格：{current_price:.2f} 元/股（最新交易日）')
    add_paragraph(document, f'• 发行日价格：{issue_date_price:.2f} 元/股（发行日{issue_date}）')
    add_paragraph(document, f'• 实际发行价：{actual_issue_price:.2f} 元/股（溢价率{(actual_issue_price/ma20_price - 1)*100:+.1f}%）')
    add_paragraph(document, '')

    # 生成溢价率对比表格数据
    premium_rate_data = []

    # 从-20%到+10%，每1%一个档
    for premium_rate in range(-20, 11, 1):
        premium_rate_decimal = premium_rate / 100.0
        calculated_issue_price = ma20_price * (1 + premium_rate_decimal)

        # 计算与发行日价格的差异
        diff_with_issue_date = calculated_issue_price - issue_date_price
        diff_with_issue_date_pct = (calculated_issue_price / issue_date_price - 1) * 100

        # 计算与当前价格的差异
        diff_with_current = calculated_issue_price - current_price
        diff_with_current_pct = (calculated_issue_price / current_price - 1) * 100

        # 判断溢价率类型
        if premium_rate < 0:
            rate_type = f"{abs(premium_rate)}%折价"
        elif premium_rate == 0:
            rate_type = "平价"
        else:
            rate_type = f"{premium_rate}%溢价"

        # 高亮实际发行价行
        if abs(calculated_issue_price - actual_issue_price) < 0.01:
            row_marker = "★"
        else:
            row_marker = ""

        premium_rate_data.append([
            f"{premium_rate:+d}%",
            rate_type,
            f"{calculated_issue_price:.2f}",
            f"{diff_with_issue_date:+.2f}",
            f"{diff_with_issue_date_pct:+.1f}%",
            f"{diff_with_current:+.2f}",
            f"{diff_with_current_pct:+.1f}%",
            row_marker
        ])

    # 添加表格
    premium_headers = ['溢价率', '类型', '发行价(元)', '与发行日价格差异(元)', '与发行日价格差异(%)', '与当前价格差异(元)', '与当前价格差异(%)', '当前方案']
    add_table_data(document, premium_headers, premium_rate_data)

    add_paragraph(document, '')
    add_paragraph(document, '附表说明：')
    add_paragraph(document, '• 溢价率：发行价相对MA20的溢价程度（负值=折价，正值=溢价）')
    add_paragraph(document, '• 类型：溢价率的文字描述（折价/平价/溢价）')
    add_paragraph(document, '• 发行价：根据MA20和溢价率计算出的发行价格 = MA20 × (1 + 溢价率)')
    add_paragraph(document, '• 与发行日价格差异：计算出的发行价与实际发行日价格之间的绝对差异和相对差异')
    add_paragraph(document, '• 与当前价格差异：计算出的发行价与当前最新交易日价格之间的绝对差异和相对差异')
    add_paragraph(document, f'• 当前方案（★）：标记实际使用的发行方案，溢价率{(actual_issue_price/ma20_price - 1)*100:+.1f}%，发行价{actual_issue_price:.2f}元')
    add_paragraph(document, '')

    add_paragraph(document, '关键观察：')
    add_paragraph(document, f'• 折价发行（溢价率<0%）：发行价低于MA20，有利于投资者获得安全边际')
    add_paragraph(document, f'• 平价发行（溢价率=0%）：发行价等于MA20价格{ma20_price:.2f}元')
    add_paragraph(document, f'• 溢价发行（溢价率>0%）：发行价高于MA20，投资者承担更高的估值风险')
    add_paragraph(document, f'• 当前方案{actual_issue_price:.2f}元相对发行日价格{issue_date_price:.2f}元{"溢价" if actual_issue_price > issue_date_price else "折价"}{abs((actual_issue_price/issue_date_price - 1)*100):.1f}%')
    add_paragraph(document, f'• 当前方案{actual_issue_price:.2f}元相对当前价格{current_price:.2f}元{"溢价" if actual_issue_price > current_price else "折价"}{abs((actual_issue_price/current_price - 1)*100):.1f}%')
    add_paragraph(document, '')

    add_paragraph(document, '使用建议：')
    add_paragraph(document, '• 参考本表评估不同溢价率方案的风险收益特征')
    add_paragraph(document, '• 对比不同溢价率下的发行价与市场价格差异，选择合适的发行方案')
    add_paragraph(document, '• 关注发行价相对当前价格的折溢价程度，评估投资价值')
    add_paragraph(document, '• 结合市场环境和项目质量，在安全边际和发行成功之间取得平衡')
    add_paragraph(document, '')