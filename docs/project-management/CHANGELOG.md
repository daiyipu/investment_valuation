# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Planned
- 性能优化与用户体验提升
- 数据可视化增强
- 高级分析功能

## [1.0.0] - 2025-02-15

### Added

#### 核心功能
- ✨ DCF绝对估值法实现
- ✨ 相对估值法（P/E、P/S、P/B、EV/EBITDA）
- ✨ 情景分析（基准/乐观/悲观）
- ✨ 压力测试模块
  - 收入冲击测试
  - 利润率压缩测试
  - WACC冲击测试
  - 蒙特卡洛模拟
  - 极端市场崩溃测试
- ✨ 敏感性分析
  - 单因素分析
  - 双因素分析
  - 龙卷风图生成

#### 前端功能
- 🎨 Vue.js 3 前端界面
- 🎨 估值输入表单
- 🎨 估值结果展示
- 🎨 ECharts数据可视化
- 🎨 申万三级行业分类选择器
- 🎨 可比公司选择与管理
- 🎨 报告页面
- 🎨 报告打印/PDF导出功能

#### 数据集成
- 📊 Tushare API集成
- 📊 A股上市公司数据获取
- 📊 申万行业分类数据
- 📊 财务指标自动提取
- 📊 交易日历API集成

#### API服务
- 🔌 FastAPI RESTful API
- 🔌 API文档自动生成（Swagger/OpenAPI）
- 🔌 数据验证与错误处理
- 🔌 CORS跨域支持
- 🔌 历史记录管理

### Fixed
- 🐛 修复估值页面computed导入问题
- 🐛 修复ValuationResult中sortedSensitivityParams调用问题
- 🐛 修复Report.vue模板div标签嵌套错误
- 🐛 修复类型错误（formatMoney/formatPercent）
- 🐛 优化sessionStorage数据保存逻辑

### Changed
- ♻️ 重构数据获取模块，使用Tushare交易日历API
- ♻️ 优化前端代码分割
- ♻️ 改进错误处理和用户提示
- ♻️ 增强类型安全（TypeScript）

### Technical Details
- **Backend**:
  - Python 3.11
  - FastAPI 0.109+
  - Pydantic数据验证
  - NumPy/Pandas数据处理

- **Frontend**:
  - Vue.js 3.4+
  - TypeScript 5.3+
  - Vite 5.0+
  - ECharts 5.4+
  - Element Plus UI组件

### Performance
- API平均响应时间: ~5秒（待优化至<2秒）
- 页面首次加载: ~5秒（待优化至<3秒）
- 并发支持: 100+用户

### Documentation
- 📝 完善API文档
- 📝 添加开发指南
- 📝 添加贡献指南
- 📝 创建项目路线图
- 📝 建立CI/CD流程

### Project Management
- 📋 制定迭代规划（v1.1-v2.0）
- 📋 建立Issue模板
- 📋 建立PR模板
- 📋 配置GitHub Actions CI/CD
- 📋 删除所有.bak备份文件

---

## [0.x] - Historical Versions

### Version History

- **v0.1**: 初始版本，基础估值算法
- **v0.2**: 添加情景分析和压力测试
- **v0.3**: 添加敏感性分析
- **v0.4**: 集成Tushare数据源
- **v0.5**: 添加FastAPI服务
- **v0.6**: 开发Vue.js前端
- **v0.7**: 完善用户界面
- **v0.8**: 添加报告生成功能
- **v0.9**: Bug修复和性能优化
- **v1.0**: 正式发布

---

## 版本说明

### 版本号规则

项目遵循 [Semantic Versioning](https://semver.org/)：

```
MAJOR.MINOR.PATCH

例如: 1.2.3
- MAJOR: 不兼容的API变更
- MINOR: 向后兼容的功能新增
- PATCH: 向后兼容的Bug修复
```

### 变更类型

- **Added**: 新增功能
- **Changed**: 现有功能的变更
- **Deprecated**: 即将移除的功能
- **Removed**: 已移除的功能
- **Fixed**: Bug修复
- **Security**: 安全性修复

---

## 链接

- [项目路线图](./PROJECT_ROADMAP.md)
- [贡献指南](./CONTRIBUTING.md)
- [当前迭代任务](./SPRINT_1.1_TASKS.md)
- [GitHub Issues](https://github.com/daiyipu/investment_valuation/issues)
- [GitHub Releases](https://github.com/daiyipu/investment_valuation/releases)
