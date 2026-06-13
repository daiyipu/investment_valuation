"""选股策略包: 波二(wave2) / 抵抗(resist) PIT 信号函数 + runner。

设计: 信号函数为纯函数(无 IO/DB), 输入序列由 runner 通过 data_loader 提供。
- indicators.py: 纯计算工具(MA/corr/swing/回撤), 可单测
- params.py: 集中阈值参数(便于调)
- wave2.py / resist.py: 信号函数
- data_loader.py: PIT 取数(≤date 截断)
- screen_strategies.py: 屏幕 runner
"""
