<template>
  <div class="report">
    <div class="header">
      <h1>📊 综合估值报告</h1>
      <p>{{ company?.name }} - {{ company?.industry }}</p>
    </div>

    <div v-if="!results" class="no-data">
      <p>暂无数据</p>
      <button @click="$router.push('/valuation')" class="btn-primary">开始估值</button>
    </div>

    <template v-else>
      <!-- 执行摘要 -->
      <div class="card summary">
        <div class="card-title">执行摘要</div>
        <div class="summary-grid">
          <div class="summary-item">
            <span class="summary-label">DCF估值</span>
            <span class="summary-value">{{ formatMoney(results.dcf?.result?.value) }}</span>
          </div>
          <div class="summary-item">
            <span class="summary-label">推荐估值</span>
            <span class="summary-value primary">{{ formatMoney(getRecommendedValue()) }}</span>
            <span class="summary-method">（中位数）</span>
          </div>
          <div class="summary-item" v-if="getUsedMethods().length > 0">
            <span class="summary-label">参考方法</span>
            <span class="summary-methods">{{ getUsedMethods().join(' + ') }}</span>
          </div>
          <div class="summary-item">
            <span class="summary-label">估值区间</span>
            <span class="summary-value">{{ getValueRange() }}</span>
          </div>
        </div>
      </div>

      <!-- 估值方法 -->
      <div class="card">
        <div class="card-title">估值方法</div>
        <div class="methods-list">
          <!-- 相对估值方法 -->
          <template v-if="results.relative?.results && Object.keys(results.relative.results).length > 0">
            <div v-for="(result, method) in results.relative.results" :key="method" class="method-item">
              <div class="method-header">
                <span class="method-name">{{ getRelativeMethodName(method) }}</span>
                <span class="method-value">{{ formatMoney(result.value) }}</span>
              </div>
              <div v-if="result.value_low && result.value_high" class="method-details">
                区间: {{ formatMoney(result.value_low) }} - {{ formatMoney(result.value_high) }}
              </div>
            </div>
          </template>

          <!-- DCF估值 -->
          <div class="method-item">
            <div class="method-header">
              <span class="method-name">DCF现金流折现</span>
              <span class="method-value">{{ formatMoney(results.dcf?.result?.value) }}</span>
            </div>
            <div class="method-details">
              WACC: {{ formatPercent(results.dcf?.result?.details?.wacc) }} |
              终值占比: {{ getTerminalPercent() }}%
            </div>
          </div>

          <!-- 相对估值 -->
          <div v-if="results.relative" class="method-item">
            <div class="method-header">
              <span class="method-name">相对估值（市场倍数法）</span>
              <span class="method-value">{{ formatMoney(results.relative.result?.value) }}</span>
            </div>
            <div class="method-details">
              <div v-if="results.relative?.result?.pe_ratio">
                P/E倍数: <strong>{{ results.relative.result.pe_ratio.toFixed(2) }}</strong> |
                估值: {{ formatMoney(results.relative.result.pe_valuation) }}
              </div>
              <div v-if="results.relative?.result?.ps_ratio">
                P/S倍数: <strong>{{ results.relative.result.ps_ratio.toFixed(2) }}</strong> |
                估值: {{ formatMoney(results.relative.result.ps_valuation) }}
              </div>
              <div v-if="results.relative?.result?.pb_ratio">
                P/B倍数: <strong>{{ results.relative.result.pb_ratio.toFixed(2) }}</strong> |
                估值: {{ formatMoney(results.relative.result.pb_valuation) }}
              </div>
              <div v-if="results.relative?.result?.ev_ebitda">
                EV/EBITDA倍数: <strong>{{ results.relative.result.ev_ebitda.toFixed(2) }}</strong> |
                估值: {{ formatMoney(results.relative.result.ev_valuation) }}
              </div>
            </div>
            <div class="method-details">
              <p v-if="!results.relative?.comparables" class="no-comparables">
                未使用可比公司
              </p>
              <p v-else class="comparables-info">
                基于 <strong>{{ results.relative.comparables?.length || 0 }}</strong> 家可比公司
                <span v-for="(comp, idx) in results.relative.comparables" :key="idx" class="comparable-company">
                  {{ comp.name }}
                </span>
              </p>
            </div>
          </div>
        </div>
      </div>

      <!-- 风险分析 -->
      <div class="card">
        <div class="card-title">风险分析</div>
        <div class="risk-section">
          <h3>情景分析</h3>
          <div class="scenario-table">
            <div v-for="(scenario, name) in getScenarios()" :key="name" class="scenario-row">
              <span>{{ name }}</span>
              <span>{{ formatMoney(scenario.valuation?.value || scenario.value) }}</span>
            </div>
          </div>
        </div>

        <div class="risk-section">
          <h3>压力测试</h3>
          <div class="stress-table-container">
            <table class="stress-table">
              <thead>
                <tr>
                  <th>压力情景</th>
                  <th>压力下估值</th>
                  <th>变化幅度</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(test, idx) in getStressTests()" :key="idx"
                    class="stress-row"
                    :class="{ 'extreme-row': test.test_name === '极端市场崩溃测试' }">
                  <td class="stress-scenario">{{ test.scenario_description }}</td>
                  <td class="stress-value">{{ formatMoney(test.stressed_value) }}</td>
                  <td class="stress-change" :class="test.change_pct < 0 ? 'negative' : 'positive'">
                    {{ test.change_pct > 0 ? '+' : '' }}{{ formatPercent(test.change_pct) }}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <div class="risk-section">
          <h3>敏感性分析</h3>
          <div class="sensitivity-table-container">
            <table class="sensitivity-table">
              <thead>
                <tr>
                  <th>参数名称</th>
                  <th>估值波动范围</th>
                  <th>影响程度</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(param, name) in sortedSensitivityParams" :key="name" class="sensitivity-row">
                  <td class="sensitivity-parameter">{{ getParameterName(name) }}</td>
                  <td class="sensitivity-value">±{{ formatMoney(param.valuation_range / 2) }}</td>
                  <td class="sensitivity-impact" :class="{ 'high-impact': name === mostSensitiveParam }">
                    {{ name === mostSensitiveParam ? '⭐ 最大影响' : '一般' }}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- 投资建议 -->
      <div class="card recommendation">
        <div class="card-title">投资建议</div>
        <div class="recommendation-content">
          <div class="recommendation-level" :class="getRecommendationLevel()">
            {{ getRecommendationText() }}
          </div>
          <div class="recommendation-reasons">
            <h4>理由：</h4>
            <ul>
              <li>DCF估值显示公司内在价值为 {{ formatMoney(results.dcf?.result?.value) }}</li>
              <li>情景分析表明估值存在一定不确定性</li>
              <li>压力测试显示最大下行风险为 {{ getMaxDownside() }}</li>
              <li>建议设置 {{ getSafetyMargin() }} 的安全边际</li>
            </ul>
          </div>
        </div>
      </div>

      <!-- 导出按钮 -->
      <div class="actions">
        <button @click="printReport" class="btn-primary">🖨️ 打印/导出PDF</button>
        <button @click="$router.push('/valuation')" class="btn-secondary">重新估值</button>
      </div>
    </template>
  </div>
</template>

<script setup lang="ts">
import { ref, onBeforeMount, computed } from 'vue'

const results = ref<any>(null)
const company = ref<any>(null)

onBeforeMount(() => {
  const data = sessionStorage.getItem('valuationResults')
  if (data) {
    const parsed = JSON.parse(data)
    results.value = parsed
    company.value = parsed.company
  }
})

const getRecommendedValue = () => {
  const values: number[] = []

  // 收集所有估值方法的结果
  if (results.value?.dcf?.result?.value) {
    values.push(results.value.dcf.result.value)
  }

  if (results.value?.relative?.results) {
    for (const result of Object.values(results.value.relative.results)) {
      if ((result as any).value) {
        values.push((result as any).value)
      }
    }
  }

  if (values.length === 0) return 0

  // 返回中位数
  values.sort((a, b) => a - b)
  const mid = Math.floor(values.length / 2)
  return values.length % 2 !== 0 ? values[mid] : (values[mid - 1] + values[mid]) / 2
}

const getUsedMethods = () => {
  const methods: string[] = []

  if (results.value?.dcf?.result?.value) {
    methods.push('DCF')
  }

  if (results.value?.relative?.results) {
    for (const method of Object.keys(results.value.relative.results)) {
      methods.push(getRelativeMethodName(method))
    }
  }

  return methods
}

const getValueRange = () => {
  const value = getRecommendedValue()
  const low = (value * 0.9 / 10000).toFixed(2)
  const high = (value * 1.1 / 10000).toFixed(2)
  return `${low} - ${high} 亿元`
}

const getTerminalPercent = () => {
  const pvTerminal = results.value?.dcf?.result?.details?.pv_terminal || 0
  const total = results.value?.dcf?.result?.value || 1
  return ((pvTerminal / total) * 100).toFixed(1)
}

const getScenarios = () => {
  const scenarios = results.value?.scenario?.results || {}
  const filtered: Record<string, any> = {}
  for (const [name, data] of Object.entries(scenarios)) {
    if (name !== 'statistics') {
      filtered[name] = data
    }
  }
  return filtered
}

const getStressTests = () => {
  const revenueTests = results.value?.stress?.report?.tests?.revenue_shock
  const extremeTest = results.value?.stress?.report?.tests?.extreme_crash

  if ((!revenueTests || revenueTests.length === 0) && !extremeTest) return []

  const allTests: any[] = []

  // 添加收入冲击测试
  if (revenueTests && revenueTests.length > 0) {
    allTests.push(...revenueTests)
  }

  // 添加极端情景（如果存在）
  if (extremeTest) {
    allTests.push(extremeTest)
  }

  // 按照严重程度排序（极端情景放最后）
  return allTests.sort((a: any, b: any) => {
    // 极端情景始终排在最后
    if (a.test_name === '极端市场崩溃测试') return 1
    if (b.test_name === '极端市场崩溃测试') return -1

    // 其他情况按照冲击幅度排序
    const getShockPct = (desc: string) => {
      const match = desc.match(/(\d+(\.\d+)?)/)
      return match ? parseFloat(match[0]) : 0
    }
    const pctA = getShockPct(a.scenario_description)
    const pctB = getShockPct(b.scenario_description)
    return pctB - pctA // 降序排列
  })
}

const getStressImpact = () => {
  const tests = getStressTests()
  if (tests.length > 0) {
    const impact = tests[0].change_pct
    return (impact * 100).toFixed(1) + '%'
  }
  return '--'
}

const getMaxDownside = () => {
  // 使用所有压力测试（包括极端情景）
  const tests = getStressTests()
  if (tests && tests.length > 0) {
    const minChange = Math.min(...tests.map((t: any) => t.change_pct))
    return (minChange * 100).toFixed(1) + '%'
  }
  return '--'
}

const sortedSensitivityParams = computed(() => {
  const params = results.value?.sensitivity?.results?.parameters
  if (!params) return {}

  // 按照valuation_range排序（从大到小）
  const entries = Object.entries(params).sort((a: any, b: any) =>
    (b[1] as any).valuation_range - (a[1] as any).valuation_range
  )

  // 转换回对象
  const sorted: Record<string, any> = {}
  for (const [name, data] of entries) {
    sorted[name] = data
  }
  return sorted
})

const mostSensitiveParam = computed(() => {
  const params = results.value?.sensitivity?.results?.parameters
  if (params) {
    const entries = Object.entries(params).sort((a: any, b: any) =>
      b[1].valuation_range - a[1].valuation_range)
    return entries[0]?.[0] || '--'
  }
  return '--'
})

const getRecommendationLevel = () => {
  return 'hold'
}

const getRecommendationText = () => {
  return '中性持有'
}

const getSafetyMargin = () => {
  return '20%'
}

const getRelativeMethodName = (method: string) => {
  const names: Record<string, string> = {
    'PE': '市盈率法 (P/E)',
    'PS': '市销率法 (P/S)',
    'PB': '市净率法 (P/B)',
    'EV_EBITDA': 'EV/EBITDA法'
  }
  return names[method] || method
}

const getParameterName = (param: string) => {
  const names: Record<string, string> = {
    'revenue_growth': '收入增长率',
    'operating_margin': '营业利润率',
    'wacc': '加权平均资本成本 (WACC)',
    'terminal_growth': '终值增长率',
    'perpetual_growth': '永续增长率'
  }
  return names[param] || param
}

const formatMoney = (value: number | string | undefined) => {
  if (value === undefined || value === null) return '--'
  const numValue = typeof value === 'string' ? parseFloat(value) : value
  if (isNaN(numValue)) return '--'
  return (numValue / 10000).toFixed(2) + ' 亿元'
}
const formatPercent = (value: number | undefined) => {
  if (value === undefined || value === null || isNaN(value)) return '--'
  return (value * 100).toFixed(2) + '%'
}

// 浏览器打印/导出PDF
const printReport = () => {
  if (!results.value) {
    alert('暂无数据可打印')
    return
  }

  // 触发浏览器打印
  window.print()
}

// 添加打印样式
const printStyles = `
  @media print {
    body * {
      visibility: hidden;
    }
    .report, .report * {
      visibility: visible;
    }
    .report {
      position: absolute;
      left: 0;
      top: 0;
      width: 100%;
      background: white;
    }
    .actions, .no-data {
      display: none;
    }
    @page {
      margin: 1cm;
    }
  }
`

// 动态添加打印样式
if (typeof document !== 'undefined') {
  const style = document.createElement('style')
  style.textContent = printStyles
  document.head.appendChild(style)
}
</script>

<style scoped>
.report {
  padding: 20px;
  max-width: 1200px;
  margin: 0 auto;
}

.header {
  text-align: center;
  margin-bottom: 30px;
  padding: 30px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  border-radius: 12px;
  color: white;
}

.header h1 {
  font-size: 2em;
  margin-bottom: 10px;
}

.no-data {
  text-align: center;
  padding: 60px 20px;
  background: white;
  border-radius: 12px;
}

.card {
  background: white;
  padding: 25px;
  border-radius: 12px;
  margin-bottom: 20px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.card.summary {
  background: linear-gradient(135deg, #f8f9ff 0%, #f0f7ff 100%);
}

.card-title {
  font-size: 1.3em;
  color: #333;
  margin-bottom: 20px;
  padding-bottom: 10px;
  border-bottom: 2px solid #667eea;
}

.summary-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 20px;
}

.summary-item {
  text-align: center;
  padding: 15px;
  background: white;
  border-radius: 8px;
}

.summary-label {
  display: block;
  color: #666;
  margin-bottom: 10px;
  font-size: 0.9em;
}

.summary-value {
  font-size: 1.5em;
  font-weight: bold;
  color: #333;
}

.summary-value.primary {
  color: #667eea;
}

.summary-method {
  display: block;
  font-size: 0.85em;
  color: #666;
  margin-top: 4px;
}

.methods-list {
  display: flex;
  flex-direction: column;
  gap: 15px;
}

.method-item {
  padding: 15px;
  background: #f8f9fa;
  border-radius: 8px;
}

.method-header {
  display: flex;
  justify-content: space-between;
  margin-bottom: 10px;
}

.method-name {
  font-weight: bold;
  color: #333;
}

.method-value {
  font-size: 1.2em;
  color: #667eea;
  font-weight: 600;
}

.method-details {
  font-size: 0.9em;
  color: #666;
}

.risk-section {
  margin-bottom: 20px;
}

.risk-section h3 {
  color: #333;
  margin-bottom: 10px;
  padding-bottom: 10px;
  border-bottom: 2px solid #667eea;
}

.scenario-table {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.scenario-row {
  display: flex;
  justify-content: space-between;
  padding: 10px 15px;
  background: #f8f9fa;
  border-radius: 4px;
}

.stress-table-container {
  overflow-x: auto;
  margin-top: 10px;
}

.stress-table {
  width: 100%;
  border-collapse: collapse;
  background: white;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.stress-table thead {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
}

.stress-table th {
  padding: 12px 16px;
  text-align: left;
  font-weight: 600;
  font-size: 0.95em;
}

.stress-table tbody tr {
  border-bottom: 1px solid #f0f0f0;
  transition: background 0.2s;
}

.stress-table tbody tr:hover {
  background: #f8f9fa;
}

.stress-table tbody tr:last-child {
  border-bottom: none;
}

.stress-table td {
  padding: 12px 16px;
  color: #333;
}

.stress-scenario {
  font-weight: 500;
  color: #555;
}

.stress-value {
  font-weight: 600;
  color: #667eea;
}

.stress-change {
  font-weight: 600;
  font-size: 0.95em;
}

.stress-change.positive {
  color: #ee6666;
}

.stress-change.negative {
  color: #91cc75;
}

.stress-row.extreme-row {
  background: linear-gradient(135deg, #fff5f5 0%, #ffe0e0 100%);
  font-weight: 500;
}

.stress-row.extreme-row:hover {
  background: linear-gradient(135deg, #ffe8e8 0%, #ffd6d6 100%);
}

.stress-row.extreme-row .stress-scenario {
  color: #cc0000;
  font-weight: 600;
}

.stress-row.extreme-row .stress-change {
  font-weight: 700;
  font-size: 1.05em;
}

.sensitivity-table-container {
  overflow-x: auto;
  margin-top: 10px;
}

.sensitivity-table {
  width: 100%;
  border-collapse: collapse;
  background: white;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.sensitivity-table thead {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
}

.sensitivity-table th {
  padding: 12px 16px;
  text-align: left;
  font-weight: 600;
  font-size: 0.95em;
}

.sensitivity-table tbody tr {
  border-bottom: 1px solid #f0f0f0;
  transition: background 0.2s;
}

.sensitivity-table tbody tr:hover {
  background: #f8f9fa;
}

.sensitivity-table tbody tr:last-child {
  border-bottom: none;
}

.sensitivity-table td {
  padding: 12px 16px;
  color: #333;
}

.sensitivity-parameter {
  font-weight: 500;
  color: #555;
}

.sensitivity-value {
  font-weight: 600;
  color: #667eea;
}

.sensitivity-impact {
  font-weight: 600;
  font-size: 0.9em;
  color: #999;
}

.sensitivity-impact.high-impact {
  color: #ee6666;
  font-weight: 700;
}

.recommendation {
  background: linear-gradient(135deg, #f0fff4 0%, #f8fff8 100%);
}

.recommendation-content {
  text-align: center;
}

.recommendation-level {
  display: inline-block;
  padding: 15px 40px;
  font-size: 1.5em;
  font-weight: bold;
  border-radius: 8px;
  margin-bottom: 20px;
}

.recommendation-level.hold {
  background: #fff3cd;
  color: #856404;
}

.recommendation-reasons {
  text-align: left;
  max-width: 800px;
  margin: 0 auto;
}

.recommendation-reasons h4 {
  color: #333;
  margin-bottom: 10px;
}

.recommendation-reasons ul {
  list-style-position: inside;
  color: #555;
  line-height: 1.8;
}

.recommendation-reasons li {
  margin: 8px 0;
  padding-left: 20px;
}

.actions {
  display: flex;
  justify-content: center;
  gap: 15px;
  margin-top: 30px;
}

.btn-primary,
.btn-secondary {
  padding: 12px 30px;
  border: none;
  border-radius: 6px;
  font-size: 16px;
  cursor: pointer;
  transition: all 0.3s;
  font-weight: 500;
}

.btn-primary {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
}

.btn-secondary {
  background: #f0f0f0;
  color: #333;
}

.btn-primary:hover,
.btn-secondary:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
}
</style>
