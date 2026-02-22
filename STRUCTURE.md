# 项目结构整理说明

## 概述

为了使项目架构更加清晰，已创建了新的目录结构，将后端代码按功能模块进行分类。

## 当前项目结构

```
investment_valuation/
├── backend/                    # 新的后端目录结构（参考/迁移目标）
│   ├── core/                  # 核心层 - 数据模型和数据库
│   │   ├── models.py          # 领域模型（Company, ValuationResult等）
│   │   └── database.py        # 数据库管理
│   │
│   ├── services/              # 服务层 - 业务逻辑
│   │   ├── absolute_valuation.py      # DCF绝对估值
│   │   ├── relative_valuation.py      # 相对估值
│   │   ├── multi_product_valuation.py # 多产品估值
│   │   ├── scenario_analysis.py       # 情景分析
│   │   ├── sensitivity_analysis.py    # 敏感性分析
│   │   ├── stress_test.py             # 压力测试
│   │   └── valuation_engine.py        # 估值引擎
│   │
│   ├── api/                   # API层 - 接口定义
│   │   ├── main.py            # FastAPI主入口（新版本）
│   │   ├── schemas.py         # Pydantic模型
│   │   ├── run.py             # 启动脚本
│   │   └── routes/            # 路由模块（待完善）
│   │
│   ├── data/                  # 数据层 - 数据获取
│   │   └── fetcher.py         # Tushare数据获取
│   │
│   └── utils/                 # 工具层 - 辅助函数
│       └── other_methods.py   # 其他估值方法
│
├── frontend-vue/              # Vue.js前端
│   ├── src/
│   │   ├── views/            # 页面组件
│   │   ├── services/         # API服务
│   │   └── assets/           # 静态资源
│   └── package.json
│
├── docs/                      # 文档
│
├── api.py                     # 当前使用的FastAPI入口（根目录）
├── models.py                  # 当前使用的模型（根目录）
├── database.py                # 当前使用的数据库（根目录）
├── absolute_valuation.py      # 当前使用的估值服务（根目录）
├── relative_valuation.py      # 当前使用的估值服务（根目录）
├── multi_product_valuation.py # 当前使用的估值服务（根目录）
├── scenario_analysis.py       # 当前使用的分析服务（根目录）
├── sensitivity_analysis.py    # 当前使用的分析服务（根目录）
├── stress_test.py             # 当前使用的测试服务（根目录）
├── valuation_engine.py        # 当前使用的引擎（根目录）
├── data_fetcher.py            # 当前使用的数据获取（根目录）
├── other_methods.py           # 当前使用的工具（根目录）
│
├── examples.py                # 示例代码
├── requirements.txt           # Python依赖
├── ARCHITECTURE.md            # 架构设计文档
├── STRUCTURE.md               # 本文件
└── README.md                  # 项目说明
```

## 模块分类说明

### 核心层 (core/)

存放数据模型和数据库管理相关代码：
- `models.py`: 定义所有领域模型（Company, Comparable, ValuationResult等）
- `database.py`: 数据库连接和ORM管理

### 服务层 (services/)

存放业务逻辑和估值算法：
- `absolute_valuation.py`: DCF现金流折现估值
- `relative_valuation.py`: P/E、P/S、P/B等相对估值方法
- `multi_product_valuation.py`: 多产品/多业务线估值
- `scenario_analysis.py`: 情景分析（基准、乐观、悲观）
- `sensitivity_analysis.py`: 敏感性分析和龙卷风图
- `stress_test.py`: 压力测试和蒙特卡洛模拟
- `valuation_engine.py`: 统一的估值引擎

### API层 (api/)

存放FastAPI接口相关代码：
- `main.py`: API主入口（路由注册、CORS配置等）
- `schemas.py`: 请求/响应的数据模型
- `routes/`: 各功能模块的路由处理

### 数据层 (data/)

存放数据获取和处理代码：
- `fetcher.py`: Tushare API数据获取

### 工具层 (utils/)

存放辅助工具函数：
- `other_methods.py`: VC法等其他估值方法

## 当前状态

- ✅ 新目录结构已创建
- ✅ 文件已复制到新目录
- ✅ 架构文档已更新
- ⚠️  当前仍使用根目录的文件（api.py等）
- 🔄 新结构的import路径需要调整（待完善）

## 运行项目

### 后端（当前方式）

```bash
# 使用根目录的 api.py
python api.py

# 或使用 uvicorn
uvicorn api:app --reload --host 0.0.0.0 --port 8000
```

### 前端

```bash
cd frontend-vue
npm run dev
```

## 未来迁移步骤

1. **更新backend/目录的import路径**
   - 使用相对导入或设置PYTHONPATH
   - 确保所有模块可以正确导入

2. **拆分api.py**
   - 将路由拆分到backend/api/routes/目录
   - 每个功能模块一个路由文件

3. **更新启动脚本**
   - 使用backend/api/main.py作为新的入口
   - 保持向后兼容

4. **清理根目录**
   - 删除根目录的旧文件
   - 更新README.md中的启动说明
