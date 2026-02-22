"""
投资估值系统 API 服务器
简单的启动脚本，使用旧的 api.py
"""
import sys
import os

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import uvicorn

if __name__ == "__main__":
    # 启动旧的 API（暂时使用根目录的 api.py）
    from api import app
    uvicorn.run(app, host="0.0.0.0", port=8000)
