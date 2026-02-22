"""
敏感性分析模块
包含单因素敏感性分析、双因素敏感性分析、龙卷风图等
"""
import numpy as np
from typing import List, Dict, Optional, Any, Tuple
from core.models import Company, ValuationResult
from services.absolute_valuation import AbsoluteValuation


class SensitivityAnalyzer:
    """敏感性分析器"""

    def __init__(self, company: Company):
        """
        初始化敏感性分析器

        Args:
            company: 目标公司
        """
        self.company = company
        self.base_valuation = AbsoluteValuation.dcf_valuation(company)
        self.base_value = self.base_valuation.value

    def one_way_sensitivity(
        self,
        param_name: str,
        param_range: Optional[Tuple[float, float]] = None,
        steps: int = 10
    ) -> Dict[str, Any]:
        """
        单因素敏感性分析

        Args:
            param_name: 参数名称（'growth_rate', 'operating_margin', 'wacc', 'terminal_growth'）
            param_range: 参数范围（min, max）
            steps: 步数

        Returns:
            敏感性分析结果
        """
        # 获取参数基准值和范围
        param_map = {
            'growth_rate': (self.company.growth_rate, 0.0, self.company.growth_rate * 2),
            'operating_margin': (self.company.operating_margin or 0.2, 0.05, 0.5),
            'wacc': (AbsoluteValuation.calculate_wacc(self.company), 0.04, 0.15),
            'terminal_growth': (self.company.terminal_growth_rate, 0.0, 0.05),
        }

        if param_name not in param_map:
            raise ValueError(f"不支持的参数: {param_name}。支持的参数: {list(param_map.keys())}")

        base_value, default_min, default_max = param_map[param_name]

        if param_range:
            min_val, max_val = param_range
        else:
            min_val, max_val = default_min, default_max

        # 生成参数值序列
        param_values = np.linspace(min_val, max_val, steps)
        valuations = []

        for val in param_values:
            try:
                if param_name == 'wacc':
                    valuation = AbsoluteValuation.dcf_valuation(self.company, wacc=val)
                elif param_name == 'terminal_growth':
                    valuation = AbsoluteValuation.dcf_valuation(
                        self.company, terminal_growth_rate=val
                    )
                else:
                    custom_assumptions = {param_name: val}
                    valuation = AbsoluteValuation.dcf_valuation(
                        self.company, custom_assumptions=custom_assumptions
                    )

                valuations.append(valuation.value)
            except:
                valuations.append(None)

        # 计算敏感性指标
        valuations = np.array(valuations)
        valid_mask = ~np.isnan(valuations)
        valid_valuations = valuations[valid_mask]
        valid_params = param_values[valid_mask]

        if len(valid_valuations) >= 2:
            # 计算弹性（估值变化百分比 / 参数变化百分比）
            pct_change_val = (valid_valuations[-1] - valid_valuations[0]) / self.base_value
            pct_change_param = (valid_params[-1] - valid_params[0]) / base_value if base_value != 0 else 1
            elasticity = pct_change_val / pct_change_param if pct_change_param != 0 else 0
        else:
            elasticity = None

        return {
            'parameter': param_name,
            'base_value': base_value,
            'param_values': param_values.tolist(),
            'valuations': valuations.tolist(),
            'min_valuation': float(np.min(valid_valuations)) if len(valid_valuations) > 0 else None,
            'max_valuation': float(np.max(valid_valuations)) if len(valid_valuations) > 0 else None,
            'valuation_range': float(np.max(valid_valuations) - np.min(valid_valuations)) if len(valid_valuations) > 0 else None,
            'elasticity': elasticity,
        }

    def two_way_sensitivity(
        self,
        param1: str,
        param2: str,
        ranges: Optional[Dict[str, Tuple[float, float]]] = None,
        steps: int = 10
    ) -> Dict[str, Any]:
        """
        双因素敏感性分析

        Args:
            param1: 第一个参数
            param2: 第二个参数
            ranges: 参数范围字典
            steps: 每个维度的步数

        Returns:
            双因素敏感性矩阵
        """
        param_map = {
            'growth_rate': (self.company.growth_rate, 0.0, self.company.growth_rate * 2),
            'operating_margin': (self.company.operating_margin or 0.2, 0.05, 0.5),
            'wacc': (AbsoluteValuation.calculate_wacc(self.company), 0.04, 0.15),
            'terminal_growth': (self.company.terminal_growth_rate, 0.0, 0.05),
        }

        if param1 not in param_map or param2 not in param_map:
            raise ValueError(f"不支持的参数组合")

        # 获取参数范围
        if ranges:
            range1 = ranges.get(param1, (param_map[param1][1], param_map[param1][2]))
            range2 = ranges.get(param2, (param_map[param2][1], param_map[param2][2]))
        else:
            range1 = (param_map[param1][1], param_map[param1][2])
            range2 = (param_map[param2][1], param_map[param2][2])

        # 生成参数网格
        param1_values = np.linspace(range1[0], range1[1], steps)
        param2_values = np.linspace(range2[0], range2[1], steps)

        # 计算估值矩阵
        valuation_matrix = np.zeros((steps, steps))

        for i, p1 in enumerate(param1_values):
            for j, p2 in enumerate(param2_values):
                try:
                    custom_assumptions = {}
                    if param1 != 'wacc' and param1 != 'terminal_growth':
                        custom_assumptions[param1] = p1
                    if param2 != 'wacc' and param2 != 'terminal_growth':
                        custom_assumptions[param2] = p2

                    kwargs = {'custom_assumptions': custom_assumptions}
                    if param1 == 'wacc' or param2 == 'wacc':
                        kwargs['wacc'] = p1 if param1 == 'wacc' else p2
                    if param1 == 'terminal_growth' or param2 == 'terminal_growth':
                        kwargs['terminal_growth_rate'] = p1 if param1 == 'terminal_growth' else p2

                    valuation = AbsoluteValuation.dcf_valuation(self.company, **kwargs)
                    valuation_matrix[i, j] = valuation.value
                except:
                    valuation_matrix[i, j] = np.nan

        return {
            'param1': param1,
            'param2': param2,
            'param1_values': param1_values.tolist(),
            'param2_values': param2_values.tolist(),
            'valuation_matrix': valuation_matrix.tolist(),
            'min_valuation': float(np.nanmin(valuation_matrix)),
            'max_valuation': float(np.nanmax(valuation_matrix)),
        }

    def tornado_chart_data(
        self,
        param_changes: Optional[Dict[str, float]] = None
    ) -> List[Dict[str, Any]]:
        """
        生成龙卷风图数据

        展示各参数对估值的影响程度排序

        Args:
            param_changes: 参数变化幅度字典

        Returns:
            龙卷风图数据列表
        """
        if param_changes is None:
            param_changes = {
                'growth_rate': 0.1,      # ±10%
                'operating_margin': 0.05,  # ±5%
                'wacc': 0.01,            # ±1%
                'terminal_growth': 0.005,  # ±0.5%
            }

        results = []

        for param_name, change in param_changes.items():
            # 增加方向
            try:
                if param_name == 'wacc':
                    val_up = AbsoluteValuation.dcf_valuation(
                        self.company, wacc=AbsoluteValuation.calculate_wacc(self.company) + change
                    )
                elif param_name == 'terminal_growth':
                    val_up = AbsoluteValuation.dcf_valuation(
                        self.company, terminal_growth_rate=self.company.terminal_growth_rate + change
                    )
                else:
                    custom_assumptions = {param_name: getattr(self.company, param_name) + change}
                    val_up = AbsoluteValuation.dcf_valuation(
                        self.company, custom_assumptions=custom_assumptions
                    )
                value_up = val_up.value
            except:
                value_up = self.base_value

            # 减少方向
            try:
                if param_name == 'wacc':
                    base_wacc = AbsoluteValuation.calculate_wacc(self.company)
                    val_down = AbsoluteValuation.dcf_valuation(
                        self.company, wacc=max(0.01, base_wacc - change)
                    )
                elif param_name == 'terminal_growth':
                    val_down = AbsoluteValuation.dcf_valuation(
                        self.company, terminal_growth_rate=max(0, self.company.terminal_growth_rate - change)
                    )
                else:
                    base_val = getattr(self.company, param_name)
                    custom_assumptions = {param_name: max(0, base_val - change)}
                    val_down = AbsoluteValuation.dcf_valuation(
                        self.company, custom_assumptions=custom_assumptions
                    )
                value_down = val_down.value
            except:
                value_down = self.base_value

            # 计算影响
            impact_up = abs(value_up - self.base_value)
            impact_down = abs(value_down - self.base_value)
            max_impact = max(impact_up, impact_down)

            results.append({
                'parameter': param_name,
                'change': change,
                'value_up': value_up,
                'value_down': value_down,
                'impact_up': impact_up,
                'impact_down': impact_down,
                'max_impact': max_impact,
                'impact_pct': max_impact / self.base_value if self.base_value > 0 else 0,
            })

        # 按影响程度排序
        results.sort(key=lambda x: x['max_impact'], reverse=True)

        return results

    def comprehensive_sensitivity_analysis(self) -> Dict[str, Any]:
        """
        综合敏感性分析

        对所有关键参数进行敏感性分析

        Returns:
            综合分析结果
        """
        results = {
            'base_valuation': self.base_value,
            'parameters': {}
        }

        # 单因素敏感性
        for param in ['growth_rate', 'operating_margin', 'wacc', 'terminal_growth']:
            try:
                param_result = self.one_way_sensitivity(param)
                results['parameters'][param] = param_result
            except Exception as e:
                print(f"参数{param}敏感性分析失败: {e}")

        # 龙卷风图数据
        results['tornado_chart'] = self.tornado_chart_data()

        return results


# ===== 辅助函数 =====

def display_sensitivity_results(results: Dict[str, Any]) -> str:
    """
    格式化显示敏感性分析结果

    Args:
        results: 敏感性分析结果

    Returns:
        格式化的报告文本
    """
    output = []
    output.append("=" * 70)
    output.append(f"{'敏感性分析报告':^68}")
    output.append("=" * 70)

    output.append(f"\n基准估值: {results['base_valuation']/10000:.2f}亿元\n")

    # 各参数敏感性
    if 'parameters' in results:
        for param_name, param_result in results['parameters'].items():
            output.append(f"【{param_name}】")
            output.append(f"  基准值: {param_result['base_value']:.2%}")
            output.append(f"  估值范围: {param_result['min_valuation']/10000:.2f} - "
                        f"{param_result['max_valuation']/10000:.2f}亿元")
            output.append(f"  波动幅度: {param_result['valuation_range']/10000:.2f}亿元")

            if param_result.get('elasticity'):
                output.append(f"  弹性系数: {param_result['elasticity']:.2f}")

            output.append("")

    # 龙卷风图
    if 'tornado_chart' in results:
        output.append("【参数影响排序（龙卷风图）】")
        for i, item in enumerate(results['tornado_chart'], 1):
            output.append(f"  {i}. {item['parameter']}: ±{item['change']:.1%} "
                        f"-> 影响{item['impact_pct']:.1%}")

    output.append("\n" + "=" * 70)

    return "\n".join(output)


def create_tornado_chart_json(tornado_data: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    创建ECharts龙卷风图配置

    Args:
        tornado_data: 龙卷风图数据

    Returns:
        ECharts配置字典
    """
    param_names = [item['parameter'] for item in tornado_data]
    up_values = [item['impact_up'] / 10000 for item in tornado_data]
    down_values = [-item['impact_down'] / 10000 for item in tornado_data]

    return {
        'title': {'text': '参数敏感性龙卷风图', 'left': 'center'},
        'tooltip': {'trigger': 'axis', 'axisPointer': {'type': 'shadow'}},
        'grid': {'left': '3%', 'right': '4%', 'bottom': '3%', 'containLabel': True},
        'xAxis': {
            'type': 'value',
            'name': '估值变化（亿元）',
        },
        'yAxis': {
            'type': 'category',
            'data': param_names,
        },
        'series': [
            {
                'name': '正向影响',
                'type': 'bar',
                'data': up_values,
                'itemStyle': {'color': '#91cc75'},
            },
            {
                'name': '负向影响',
                'type': 'bar',
                'data': down_values,
                'itemStyle': {'color': '#ee6666'},
            },
        ],
    }


# ===== 使用示例 =====

if __name__ == "__main__":
    from models import Company, CompanyStage

    # 创建测试公司
    company = Company(
        name="测试股份有限公司",
        industry="软件服务",
        stage=CompanyStage.GROWTH,
        revenue=50000,
        net_income=8000,
        ebitda=12000,
        net_assets=20000,
        growth_rate=0.25,
        operating_margin=0.25,
        beta=1.2,
        terminal_growth_rate=0.025,
    )

    # 创建敏感性分析器
    analyzer = SensitivityAnalyzer(company)

    # 综合敏感性分析
    print("执行综合敏感性分析...")
    results = analyzer.comprehensive_sensitivity_analysis()

    # 显示结果
    print(display_sensitivity_results(results))

    # 双因素敏感性示例
    print("\n【双因素敏感性分析：增长率 vs WACC】")
    two_way = analyzer.two_way_sensitivity('growth_rate', 'wacc', steps=5)
    print(f"  最小估值: {two_way['min_valuation']/10000:.2f}亿元")
    print(f"  最大估值: {two_way['max_valuation']/10000:.2f}亿元")
    print(f"  估值波动范围: {(two_way['max_valuation'] - two_way['min_valuation'])/10000:.2f}亿元")

    # 龙卷风图JSON配置（可用于前端ECharts）
    tornado_json = create_tornado_chart_json(results['tornado_chart'])
    print(f"\n【龙卷风图ECharts配置】")
    print(f"  参数数量: {len(tornado_json['yAxis']['data'])}")
    print(f"  系列数量: {len(tornado_json['series'])}")
