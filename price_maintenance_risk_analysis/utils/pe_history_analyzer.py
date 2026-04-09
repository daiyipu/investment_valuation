# -*- coding: utf-8 -*-
"""
PE历史分位数分析工具

功能：
- 从tushare获取个股和申万行业的历史PE数据
- 计算历史分位数
- 生成PE趋势对比图
"""

import tushare as ts
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from datetime import datetime, timedelta
import os


class PEHistoryAnalyzer:
    """PE历史分位数分析器"""

    def __init__(self, ts_token=None):
        """
        初始化分析器

        参数:
            ts_token: tushare token，如果不提供则从环境变量获取
        """
        if ts_token is None:
            ts_token = os.environ.get('TUSHARE_TOKEN')

        if not ts_token:
            raise ValueError("TUSHARE_TOKEN未设置")

        self.pro = ts.pro_api(ts_token)

        # 设置中文字体
        font_path = self._find_chinese_font()
        if font_path:
            self.font_prop = fm.FontProperties(fname=font_path)
            plt.rcParams['font.sans-serif'] = [self.font_prop.get_name()]
        else:
            self.font_prop = None
            plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti']

        plt.rcParams['axes.unicode_minus'] = False

    def _find_chinese_font(self):
        """查找中文字体"""
        possible_fonts = [
            '/System/Library/Fonts/STHeiti Light.ttc',  # macOS
            '/System/Library/Fonts/PingFang.ttc',  # macOS
            'C:/Windows/Fonts/simhei.ttf',  # Windows
            'C:/Windows/Fonts/msyh.ttc',  # Windows
            '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',  # Linux
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',  # Linux
        ]

        for font_path in possible_fonts:
            if os.path.exists(font_path):
                return font_path
        return None

    def get_stock_pe_history(self, stock_code, days=1825, max_retries=3):
        """
        获取个股历史PE数据（带重试机制）

        参数:
            stock_code: 股票代码，如 '300735.SZ'
            days: 历史自然日天数，默认1825天（5年）
                  实际交易日数据会少于自然日（约60-65%）
            max_retries: 最大重试次数，默认3次

        返回:
            DataFrame: 包含trade_date, pe_ttm, pb, ps_ttm等列
        """
        end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y%m%d')

        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    print(f"  第{attempt + 1}次尝试获取个股PE数据...")

                df = self.pro.daily_basic(
                    ts_code=stock_code,
                    start_date=start_date,
                    end_date=end_date
                )

                if df is None or df.empty:
                    if attempt < max_retries - 1:
                        print(f"  第{attempt + 1}次尝试未获取到数据，{2}秒后重试...")
                        import time
                        time.sleep(2)
                        continue
                    else:
                        print(f"⚠️ 未获取到{stock_code}的历史PE数据")
                        print(f"   可能原因：该股票上市时间不足5年，或API接口返回空数据")
                        print(f"   建议：减少历史天数参数（如days=730获取2年数据）")
                        return None

                # 过滤掉PE为空或异常的数据
                df = df[df['pe_ttm'] > 0].copy()
                df = df.sort_values('trade_date').reset_index(drop=True)

                if len(df) > 0:
                    return df
                else:
                    if attempt < max_retries - 1:
                        print(f"  第{attempt + 1}次尝试获取的数据过滤后为空，{2}秒后重试...")
                        import time
                        time.sleep(2)
                        continue
                    else:
                        print(f"⚠️ {stock_code}的历史PE数据过滤后为空")
                        return None

            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"  第{attempt + 1}次尝试获取个股PE数据失败: {e}，{2}秒后重试...")
                    import time
                    time.sleep(2)
                    continue
                else:
                    print(f"❌ 获取{stock_code}历史PE数据失败: {e}")
                    return None

    def get_industry_pe_history(self, stock_code, days=1825, max_retries=3):
        """
        获取申万行业历史PE数据（带重试机制）

        参数:
            stock_code: 股票代码，如 '300735.SZ'
            days: 历史自然日天数，默认1825天（5年）
                  实际交易日数据会少于自然日（约60-65%）
            max_retries: 最大重试次数，默认3次

        返回:
            tuple: (行业名称, 行业代码, DataFrame或None)
        """
        end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y%m%d')

        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    print(f"  第{attempt + 1}次尝试获取行业PE数据...")

                # 1. 获取股票所在的申万行业
                df_member = self.pro.index_member(ts_code=stock_code)

                if df_member.empty:
                    if attempt < max_retries - 1:
                        print(f"  第{attempt + 1}次尝试未获取到行业成分信息，{2}秒后重试...")
                        import time
                        time.sleep(2)
                        continue
                    else:
                        print(f" 未获取到{stock_code}的行业成分信息")
                        return None, None, None

                # 过滤申万行业（优先三级，其次二级，最后一级）
                df_sw = df_member[df_member['index_code'].str.contains('.SI')]

                # 找当前有效的行业（out_date为空的）
                df_current = df_sw[df_sw['out_date'].isna()]

                if df_current.empty:
                    # 如果没有当前有效的，使用最近加入的
                    df_current = df_sw.sort_values('in_date', ascending=False).head(1)

                # 按优先级查找：三级 > 二级 > 一级
                index_code = None
                index_level = None

                # 申万三级：85开头
                l3_codes = df_current[df_current['index_code'].str.match(r'^85\d{4}\.SI$')]
                if not l3_codes.empty:
                    index_code = l3_codes.iloc[0]['index_code']
                    index_level = 'L3'
                # 申万二级：80开头
                else:
                    l2_codes = df_current[df_current['index_code'].str.match(r'^80\d{4}\.SI$')]
                    if not l2_codes.empty:
                        index_code = l2_codes.iloc[0]['index_code']
                        index_level = 'L2'
                    # 申万一级：80开头（前两位）
                    else:
                        # 尝试获取行业分类信息
                        df_classify = self.pro.index_classify(level='L1', src='SW2021')
                        # 这里简化处理，使用第一个申万指数
                        if not df_current.empty:
                            index_code = df_current.iloc[0]['index_code']
                            index_level = 'L1'

                if index_code is None:
                    print(f" 未找到{stock_code}的申万行业代码")
                    return None, None, None

                # 获取行业名称
                df_classify = self.pro.index_classify(level=index_level, src='SW2021')
                industry_info = df_classify[df_classify['index_code'] == index_code]
                industry_name = industry_info.iloc[0]['industry_name'] if not industry_info.empty else index_code

                # 2. 获取行业指数历史PE数据
                # 使用sw_daily接口获取申万行业日线行情（包含PE、PB等估值指标）
                try:
                    df_index = self.pro.sw_daily(
                        ts_code=index_code,
                        start_date=start_date,
                        end_date=end_date
                    )

                    if df_index is None or df_index.empty:
                        print(f" {industry_name}({index_code})的历史PE数据不可用")
                        print(f"   原因：sw_daily接口返回空数据")
                        print(f"   解决方案：将使用个股PE历史数据生成趋势图")
                        return industry_name, index_code, None

                    # sw_daily返回的是pe列（不是pe_ttm），需要重命名以保持一致性
                    if 'pe' in df_index.columns:
                        df_index = df_index.rename(columns={'pe': 'pe_ttm'})
                    else:
                        print(f" {industry_name}({index_code})数据中不包含PE列")
                        return industry_name, index_code, None

                except Exception as api_error:
                    print(f" 调用tushare sw_daily接口失败")
                    print(f"   原因：{api_error}")
                    print(f"   解决方案：将使用个股PE历史数据生成趋势图")
                    return industry_name, index_code, None

                # 过滤掉PE为空或异常的数据
                df_index = df_index[df_index['pe_ttm'] > 0].copy()
                df_index = df_index.sort_values('trade_date').reset_index(drop=True)

                return industry_name, index_code, df_index

            except Exception as e:
                print(f" 获取{stock_code}行业历史PE数据失败: {e}")
                return None, None, None

    def calculate_pe_percentiles(self, df):
        """
        计算PE的历史分位数

        参数:
            df: 包含pe_ttm列的DataFrame

        返回:
            DataFrame: 新增percentile列
        """
        if df is None or df.empty:
            return None

        df = df.copy()

        # 计算累计分位数（到每个交易日为止的历史分位数）
        percentiles = []
        for i in range(len(df)):
            # 使用从开始到当前的数据计算分位数
            historical_data = df['pe_ttm'].iloc[:i+1]
            current_pe = df['pe_ttm'].iloc[i]
            percentile = (historical_data < current_pe).sum() / len(historical_data) * 100
            percentiles.append(percentile)

        df['pe_percentile'] = percentiles

        return df

    def generate_pe_trend_chart(self, stock_code, stock_data, industry_name, industry_data, save_path):
        """
        生成PE趋势对比图

        参数:
            stock_code: 股票代码
            stock_data: 个股历史PE数据
            industry_name: 行业名称
            industry_data: 行业历史PE数据
            save_path: 图表保存路径
        """
        if stock_data is None or industry_data is None:
            print("⚠️ 数据缺失，无法生成图表")
            return None

        fig, axes = plt.subplots(2, 2, figsize=(16, 12))

        # 转换日期格式
        stock_data['trade_date_dt'] = pd.to_datetime(stock_data['trade_date'])
        industry_data['trade_date_dt'] = pd.to_datetime(industry_data['trade_date'])

        # 1. PE走势对比（左上）
        ax1 = axes[0, 0]
        ax1.plot(stock_data['trade_date_dt'], stock_data['pe_ttm'],
                label=f'{stock_code} PE-TTM', linewidth=2, color='#2E86AB')
        ax1.plot(industry_data['trade_date_dt'], industry_data['pe_ttm'],
                label=f'{industry_name} PE-TTM', linewidth=2, color='#A23B72')

        # 标注当前值
        current_stock_pe = stock_data.iloc[-1]['pe_ttm']
        current_industry_pe = industry_data.iloc[-1]['pe_ttm']
        ax1.axhline(y=current_stock_pe, color='#2E86AB', linestyle='--', alpha=0.5)
        ax1.axhline(y=current_industry_pe, color='#A23B72', linestyle='--', alpha=0.5)

        ax1.set_xlabel('日期', fontsize=12, fontproperties=self.font_prop)
        ax1.set_ylabel('PE-TTM', fontsize=12, fontproperties=self.font_prop)
        ax1.set_title('PE走势对比', fontsize=14, fontweight='bold', fontproperties=self.font_prop)
        ax1.legend(loc='best', prop=self.font_prop)
        ax1.grid(True, alpha=0.3)
        ax1.tick_params(axis='x', rotation=45)

        # 添加数值标注
        ax1.text(0.02, 0.98, f'当前个股PE: {current_stock_pe:.2f}',
                transform=ax1.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='#2E86AB', alpha=0.2),
                fontsize=10, fontproperties=self.font_prop)
        ax1.text(0.02, 0.90, f'当前行业PE: {current_industry_pe:.2f}',
                transform=ax1.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='#A23B72', alpha=0.2),
                fontsize=10, fontproperties=self.font_prop)

        # 2. PE相对位置（个股/行业）（右上）
        ax2 = axes[0, 1]

        # 合并数据计算相对位置
        merged = pd.merge(
            stock_data[['trade_date_dt', 'pe_ttm']],
            industry_data[['trade_date_dt', 'pe_ttm']],
            on='trade_date_dt',
            suffixes=('_stock', '_industry')
        )
        merged['pe_ratio'] = merged['pe_ttm_stock'] / merged['pe_ttm_industry']

        ax2.plot(merged['trade_date_dt'], merged['pe_ratio'],
                linewidth=2, color='#F18F01')
        ax2.axhline(y=1.0, color='black', linestyle='--', alpha=0.5, label='行业平均')
        ax2.axhline(y=merged['pe_ratio'].iloc[-1], color='red', linestyle=':', alpha=0.7)

        ax2.set_xlabel('日期', fontsize=12, fontproperties=self.font_prop)
        ax2.set_ylabel('个股PE / 行业PE', fontsize=12, fontproperties=self.font_prop)
        ax2.set_title('PE相对位置（个股/行业）', fontsize=14, fontweight='bold', fontproperties=self.font_prop)
        ax2.legend(loc='best', prop=self.font_prop)
        ax2.grid(True, alpha=0.3)
        ax2.tick_params(axis='x', rotation=45)

        # 添加当前值标注
        current_ratio = merged['pe_ratio'].iloc[-1]
        ax2.text(0.02, 0.95, f'当前相对位置: {current_ratio:.2f}x',
                transform=ax2.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5),
                fontsize=11, fontproperties=self.font_prop)

        # 3. 个股PE历史分位数（左下）
        ax3 = axes[1, 0]

        # 计算个股历史分位数
        stock_data = self.calculate_pe_percentiles(stock_data)

        ax3.plot(stock_data['trade_date_dt'], stock_data['pe_percentile'],
                linewidth=2, color='#C73E1D')
        ax3.fill_between(stock_data['trade_date_dt'], 0, stock_data['pe_percentile'], alpha=0.3, color='#C73E1D')
        ax3.axhline(y=50, color='black', linestyle='--', alpha=0.5, label='历史中位数')
        ax3.axhline(y=25, color='orange', linestyle=':', alpha=0.5, label='25分位数')
        ax3.axhline(y=75, color='orange', linestyle=':', alpha=0.5, label='75分位数')

        current_percentile = stock_data.iloc[-1]['pe_percentile']
        ax3.axhline(y=current_percentile, color='red', linestyle='-', linewidth=2, label=f'当前({current_percentile:.1f}%)')

        ax3.set_xlabel('日期', fontsize=12, fontproperties=self.font_prop)
        ax3.set_ylabel('历史分位数 (%)', fontsize=12, fontproperties=self.font_prop)
        ax3.set_title(f'{stock_code} PE历史分位数趋势', fontsize=14, fontweight='bold', fontproperties=self.font_prop)
        ax3.set_ylim(0, 100)
        ax3.legend(loc='best', prop=self.font_prop, fontsize=9)
        ax3.grid(True, alpha=0.3)
        ax3.tick_params(axis='x', rotation=45)

        # 添加统计信息
        current_pe_value = stock_data.iloc[-1]['pe_ttm']
        ax3.text(0.02, 0.95, f'当前PE: {current_pe_value:.2f}',
                transform=ax3.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='#C73E1D', alpha=0.2),
                fontsize=10, fontproperties=self.font_prop)

        # 4. 行业PE历史分位数（右下）
        ax4 = axes[1, 1]

        # 计算行业历史分位数
        industry_data = self.calculate_pe_percentiles(industry_data)

        ax4.plot(industry_data['trade_date_dt'], industry_data['pe_percentile'],
                linewidth=2, color='#6A994E')
        ax4.fill_between(industry_data['trade_date_dt'], 0, industry_data['pe_percentile'], alpha=0.3, color='#6A994E')
        ax4.axhline(y=50, color='black', linestyle='--', alpha=0.5, label='历史中位数')
        ax4.axhline(y=25, color='orange', linestyle=':', alpha=0.5, label='25分位数')
        ax4.axhline(y=75, color='orange', linestyle=':', alpha=0.5, label='75分位数')

        current_industry_percentile = industry_data.iloc[-1]['pe_percentile']
        ax4.axhline(y=current_industry_percentile, color='red', linestyle='-', linewidth=2, label=f'当前({current_industry_percentile:.1f}%)')

        ax4.set_xlabel('日期', fontsize=12, fontproperties=self.font_prop)
        ax4.set_ylabel('历史分位数 (%)', fontsize=12, fontproperties=self.font_prop)
        ax4.set_title(f'{industry_name} PE历史分位数趋势', fontsize=14, fontweight='bold', fontproperties=self.font_prop)
        ax4.set_ylim(0, 100)
        ax4.legend(loc='best', prop=self.font_prop, fontsize=9)
        ax4.grid(True, alpha=0.3)
        ax4.tick_params(axis='x', rotation=45)

        # 添加统计信息
        industry_pe_value = industry_data.iloc[-1]['pe_ttm']
        ax4.text(0.02, 0.95, f'当前PE: {industry_pe_value:.2f}',
                transform=ax4.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='#6A994E', alpha=0.2),
                fontsize=10, fontproperties=self.font_prop)

        plt.suptitle('PE历史分位数趋势分析', fontsize=16, fontweight='bold', fontproperties=self.font_prop)
        plt.tight_layout()

        # 保存图表
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        plt.savefig(save_path, dpi=150, bbox_inches='tight')
        plt.close()

        print(f"✅ PE趋势图已保存: {save_path}")
        return save_path


# 使用示例
if __name__ == '__main__':
    analyzer = PEHistoryAnalyzer()

    stock_code = '300735.SZ'

    # 获取数据
    stock_data = analyzer.get_stock_pe_history(stock_code, days=730)
    industry_name, industry_code, industry_data = analyzer.get_industry_pe_history(stock_code, days=730)

    if stock_data is not None and industry_data is not None:
        print(f"\n=== 数据统计 ===")
        print(f"个股PE: 当前={stock_data.iloc[-1]['pe_ttm']:.2f}, "
              f"最小={stock_data['pe_ttm'].min():.2f}, "
              f"最大={stock_data['pe_ttm'].max():.2f}")
        print(f"行业PE: 当前={industry_data.iloc[-1]['pe_ttm']:.2f}, "
              f"最小={industry_data['pe_ttm'].min():.2f}, "
              f"最大={industry_data['pe_ttm'].max():.2f}")

        # 生成图表
        save_path = f'../images/pe_trend_chart_{stock_code.replace(".", "_")}.png'
        analyzer.generate_pe_trend_chart(
            stock_code, stock_data,
            industry_name, industry_data,
            save_path
        )
