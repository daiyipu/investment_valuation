"""统一解析 Tushare token(代码内不再硬编码)。

加载顺序:
  1. 环境变量 TUSHARE_TOKEN(env/CLI 是可信输入)
  2. 本地 .tushare_token 文件(项目根, 已 gitignore, 仅供开发机兜底)
两者都没有则抛 RuntimeError(显式失败, 优于静默用错 token)。

调用方约定:
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
tushare 会读取该环境变量, 故保持 setdefault 形态即可兼容现有 pro_api()/pro_bar 调用。
"""
import os
from pathlib import Path

_PKG_ROOT = Path(__file__).resolve().parent          # price_maintenance_risk_analysis/
_LOCAL_TOKEN = _PKG_ROOT / '.tushare_token'


def resolve_tushare_token() -> str:
    token = os.environ.get('TUSHARE_TOKEN')
    if token:
        return token
    if _LOCAL_TOKEN.exists():
        val = _LOCAL_TOKEN.read_text(encoding='utf-8').strip()
        if val:
            return val
    raise RuntimeError(
        'TUSHARE_TOKEN 未配置。请 `export TUSHARE_TOKEN=...`，'
        f'或在 {_LOCAL_TOKEN} 写入 token（该文件已 gitignore）。'
    )
