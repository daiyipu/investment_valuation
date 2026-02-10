import axios from 'axios'

const API_BASE = 'http://localhost:8000'

// 创建axios实例
const apiClient = axios.create({
  baseURL: API_BASE,
  headers: {
    'Content-Type': 'application/json'
  }
})

// 估值相关API
export const valuationAPI = {
  // 绝对估值 (DCF)
  dcf: (company: any) => apiClient.post('/api/valuation/absolute', company),

  // 相对估值
  relative: (company: any, comparables: any[], methods?: string[]) =>
    apiClient.post('/api/valuation/relative', {
      company,
      comparables,
      methods
    }),

  // 多方法交叉验证
  compare: (company: any, comparables?: any[]) =>
    apiClient.post('/api/valuation/compare', {
      company,
      comparables
    })
}

// 情景分析API
export const scenarioAPI = {
  analyze: (company: any, scenarios?: any[]) =>
    apiClient.post('/api/scenario/analyze', {
      company,
      scenarios
    })
}

// 压力测试API
export const stressTestAPI = {
  // 收入冲击测试
  revenueShock: (company: any, shocks?: number[]) =>
    apiClient.post('/api/stress-test/revenue', {
      company,
      shocks
    }),

  // 蒙特卡洛模拟
  monteCarlo: (company: any, iterations: number = 1000) =>
    apiClient.post('/api/stress-test/monte-carlo', {
      company,
      iterations
    }),

  // 综合压力测试
  full: (company: any) =>
    apiClient.post('/api/stress-test/full', company)
}

// 敏感性分析API
export const sensitivityAPI = {
  // 单因素敏感性分析
  oneWay: (company: any, paramName: string, paramRange?: number[]) =>
    apiClient.post('/api/sensitivity/one-way', {
      company,
      param_name: paramName,
      param_range: paramRange
    }),

  // 龙卷风图数据
  tornado: (company: any, paramChanges?: Record<string, number>) =>
    apiClient.post('/api/sensitivity/tornado', {
      company,
      param_changes: paramChanges
    }),

  // 综合敏感性分析
  comprehensive: (company: any) =>
    apiClient.post('/api/sensitivity/comprehensive', company)
}

// 数据获取API
export const dataAPI = {
  // 配置Tushare Token
  configureToken: (token: string) =>
    apiClient.post('/api/data/tushare/configure', null, {
      params: { token }
    }),

  // 获取可比公司
  getComparables: (industry: string, marketCapMin?: number, marketCapMax?: number, limit: number = 20) =>
    apiClient.get(`/api/data/comparable/${industry}`, {
      params: { market_cap_min: marketCapMin, market_cap_max: marketCapMax, limit }
    }),

  // 获取行业估值倍数
  getIndustryMultiples: (industry: string, method: string = 'median') =>
    apiClient.get(`/api/data/industry-multiples/${industry}`, {
      params: { method }
    }),

  // 搜索公司
  search: (keywords: string[], limit: number = 10) =>
    apiClient.post('/api/data/search', { keywords, limit })
}

export default apiClient
