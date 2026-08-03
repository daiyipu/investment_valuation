<template>
  <div class="sensitivity-analysis">
    <div class="header">
      <h1>📊 敏感性分析</h1>
      <p>评估各参数对估值的影响程度</p>
    </div>

    <!-- 高级配置 -->
    <div class="card">
      <div class="section-title">
        ⚙️ 参数配置
        <button @click="showAdvancedConfig = !showAdvancedConfig" class="btn-toggle">
          {{ showAdvancedConfig ? '收起 ▲' : '展开 ▼' }}
        </button>
      </div>
      <div v-if="showAdvancedConfig" class="advanced-config-content">
        <div class="config-section">
          <h4 class="config-title">参数变化幅度设置</h4>
          <div class="params-config-grid">
            <div class="config-item">
              <label>收入增长率变化</label>
              <div class="input-group">
                <input v-model.number="paramChanges.growth_rate" type="number" step="1" min="1" max="50" />
                <span class="input-unit">%</span>
              </div>
            </div>
            <div class="config-item">
              <label>营业利润率变化</label>
              <div class="input-group">
                <input v-model.number="paramChanges.operating_margin" type="number" step="1" min="1" max="30" />
                <span class="input-unit">%</span>
              </div>
            </div>
            <div class="config-item">
              <label>WACC变化</label>
              <div class="input-group">
                <input v-model.number="paramChanges.wacc" type="number" step="0.5" min="0.5" max="5" />
                <span class="input-unit">%</span>
              </div>
            </div>
            <div class="config-item">
              <label>终值增长率变化</label>
              <div class="input-group">
                <input v-model.number="paramChanges.terminal_growth" type="number" step="0.1" min="0.1" max="2" />
                <span class="input-unit">%</span>
              </div>
            </div>
          </div>
          <div class="config-hint">
            💡 设置各参数在敏感性分析中的变化幅度，值越大表示测试范围越广
          </div>
          <div class="config-actions">
            <button @click="runSensitivityAnalysis" class="btn-primary" :disabled="loading">
              {{ loading ? '分析中...' : '重新运行分析' }}
            </button>
          </div>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">参数敏感性排序（龙卷风图）</div>
      <div ref="tornadoChart" class="chart"></div>
    </div>

    <div class="card">
      <div class="card-title">参数详情</div>
      <div class="params-grid">
        <div v-for="(param, name) in sensitivityParams" :key="name" class="param-card">
          <div class="param-header">{{ name }}</div>
          <div class="param-value">估值影响: ±{{ formatMoney(param.valuation_range / 2) }}</div>
          <div class="param-detail">
            基准值: {{ formatParamValue(name, param.base_value) }}
          </div>
          <div class="param-detail">
            影响: {{ (param.impact_percentage * 100).toFixed(1) }}%
          </div>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">关键发现</div>
      <div class="insights">
        <div class="insight-item">
          <span class="insight-icon">🔍</span>
          <div class="insight-content">
            <div class="insight-title">最敏感参数</div>
            <div class="insight-text">{{ mostSensitiveParam }}对估值影响最大</div>
          </div>
        </div>
        <div class="insight-item">
          <span class="insight-icon">💡</span>
          <div class="insight-content">
            <div class="insight-title">建议</div>
            <div class="insight-text">应重点关注最敏感参数的准确性和合理性</div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, computed } from 'vue'
import * as echarts from 'echarts'
import { sensitivityAPI } from '../services/api'

const sensitivityParams = ref<Record<string, any>>({})
const tornadoChart = ref<HTMLElement>()
const showAdvancedConfig = ref(false)
const loading = ref(false)

const paramChanges = ref({
  growth_rate: 10,      // ±10%
  operating_margin: 5,  // ±5%
  wacc: 1,              // ±1%
  terminal_growth: 0.5  // ±0.5%
})

const mostSensitiveParam = computed(() => {
  const entries = Object.entries(sensitivityParams.value)
  if (entries.length === 0) return '--'
  return entries.sort((a, b) => b[1].valuation_range - a[1].valuation_range)[0][0]
})

onMounted(async () => {
  const data = sessionStorage.getItem('valuationResults')
  if (data) {
    const parsed = JSON.parse(data)

    if (parsed.sensitivity?.results?.parameters) {
      sensitivityParams.value = parsed.sensitivity.results.parameters
      initTornadoChart()
    }
  }
})

const runSensitivityAnalysis = async () => {
  loading.value = true
  try {
    const data = sessionStorage.getItem('valuationResults')
    if (!data) {
      alert('请先进行估值分析')
      return
    }

    const parsed = JSON.parse(data)

    // 检查是否为多产品模式
    const isMultiProduct = parsed.valuationMode === 'multi' && parsed.products && parsed.products.length > 0
    if (isMultiProduct) {
      alert('多产品估值模式暂不支持敏感性分析。\n请使用单产品估值模式进行敏感性分析。')
      loading.value = false
      return
    }

    const company = parsed.company

    // 确保包含所有必填字段（支持历史记录格式）
    const companyForAPI = {
      name: company.name || '',
      industry: company.industry || '',
      stage: company.stage || '成长期',
      revenue: company.revenue || 0,
      net_income: company.net_income ?? (company.revenue ? company.revenue * 0.15 : 0),
      ebitda: company.ebitda,
      gross_profit: company.gross_profit,
      operating_cash_flow: company.operating_cash_flow,
      total_assets: company.total_assets,
      net_assets: company.net_assets,
      total_debt: company.total_debt ?? 0,
      cash_and_equivalents: company.cash_and_equivalents ?? 0,
      growth_rate: company.growth_rate ?? 0.15,
      margin: company.margin,
      operating_margin: company.operating_margin,
      tax_rate: company.tax_rate ?? 0.25,
      beta: company.beta ?? 1.0,
      risk_free_rate: company.risk_free_rate ?? 0.03,
      market_risk_premium: company.market_risk_premium ?? 0.07,
      cost_of_debt: company.cost_of_debt ?? 0.05,
      target_debt_ratio: company.target_debt_ratio ?? 0.3,
      terminal_growth_rate: company.terminal_growth_rate ?? 0.025
    }

    // 构建参数变化字典
    const paramChangesDict: Record<string, number> = {
      growth_rate: paramChanges.value.growth_rate / 100,
      operating_margin: paramChanges.value.operating_margin / 100,
      wacc: paramChanges.value.wacc / 100,
      terminal_growth: paramChanges.value.terminal_growth / 100
    }

    // 调用龙卷风图API
    const response = await sensitivityAPI.tornado(companyForAPI, paramChangesDict)

    // 后端返回的是数组，需要转换为对象格式
    const resultArray = response.data.result || []
    sensitivityParams.value = {}
    resultArray.forEach((item: any) => {
      sensitivityParams.value[item.parameter] = {
        valuation_range: item.max_impact * 2,  // 估值波动范围
        base_value: 0,  // 基准值（暂时设为0）
        impact_percentage: item.impact_pct
      }
    })

    // 保存到 sessionStorage
    sessionStorage.setItem('valuationResults', JSON.stringify({
      ...parsed,
      sensitivity: {
        results: {
          parameters: sensitivityParams.value
        }
      }
    }))

    // 重新初始化图表
    initTornadoChart()
  } catch (error) {
    console.error('敏感性分析失败:', error)
    alert('敏感性分析失败，请检查参数设置')
  } finally {
    loading.value = false
  }
}

const initTornadoChart = () => {
  if (!tornadoChart.value) return

  const chart = echarts.init(tornadoChart.value)
  const params: string[] = []
  const impacts: number[] = []

  for (const [paramName, paramData] of Object.entries(sensitivityParams.value)) {
    if (paramData?.valuation_range) {
      params.push(paramName)
      impacts.push(parseFloat((paramData.valuation_range / 2 / 10000).toFixed(2)))
    }
  }

  // 按影响程度排序
  const sorted = params.map((p, i) => ({ name: p, impact: impacts[i] ?? 0 }))
    .sort((a, b) => b.impact - a.impact)

  chart.setOption({
    title: { text: '参数敏感性分析', left: 'center' },
    tooltip: { trigger: 'axis', formatter: '{b}: ±{c} 亿元' },
    xAxis: { type: 'value', name: '估值波动（亿元）' },
    yAxis: {
      type: 'category',
      data: sorted.map(s => s.name).reverse()
    },
    series: [{
      type: 'bar',
      data: sorted.map(s => s.impact).reverse(),
      itemStyle: { color: '#667eea' }
    }]
  })
}

const formatMoney = (value: number) => (value / 10000).toFixed(2) + ' 亿元'
const formatParamValue = (name: string, value: number) => {
  if (name.includes('率') || name.includes('率')) {
    return (value * 100).toFixed(1) + '%'
  }
  return value.toFixed(2)
}
</script>

<style scoped>
.sensitivity-analysis {
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

.card {
  background: white;
  padding: 25px;
  border-radius: 12px;
  margin-bottom: 20px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.card-title {
  font-size: 1.3em;
  color: #333;
  margin-bottom: 20px;
  padding-bottom: 10px;
  border-bottom: 2px solid #667eea;
}

.chart {
  height: 450px;
}

.params-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
  gap: 15px;
}

.param-card {
  background: #f8f9fa;
  padding: 20px;
  border-radius: 8px;
  text-align: center;
}

.param-header {
  font-size: 1.1em;
  font-weight: bold;
  color: #333;
  margin-bottom: 12px;
}

.param-value {
  font-size: 1.3em;
  color: #667eea;
  font-weight: 600;
  margin-bottom: 10px;
}

.param-detail {
  font-size: 0.9em;
  color: #666;
  margin-top: 5px;
}

.insights {
  display: grid;
  gap: 15px;
}

.insight-item {
  display: flex;
  align-items: center;
  gap: 15px;
  padding: 15px;
  background: #f0f7ff;
  border-radius: 8px;
}

.insight-icon {
  font-size: 2em;
}

.insight-title {
  font-weight: bold;
  color: #333;
  margin-bottom: 5px;
}

.insight-text {
  color: #666;
  font-size: 0.95em;
}

/* 高级配置样式 */
.section-title {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin: 0 0 15px 0;
  font-size: 1.1em;
  color: #333;
  font-weight: 600;
  padding-bottom: 10px;
  border-bottom: 1px solid #e0e0e0;
}

.btn-toggle {
  background: transparent;
  color: #667eea;
  border: 1px solid #667eea;
  padding: 4px 12px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 0.85em;
  transition: all 0.3s;
}

.btn-toggle:hover {
  background: #667eea;
  color: white;
}

.advanced-config-content {
  margin-top: 20px;
}

.config-section {
  margin-bottom: 25px;
}

.config-section:last-child {
  margin-bottom: 0;
}

.config-title {
  margin: 0 0 15px 0;
  font-size: 1em;
  color: #333;
  font-weight: 600;
}

.params-config-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 15px;
  margin-bottom: 15px;
}

.config-item {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.config-item label {
  font-size: 0.9em;
  color: #555;
  font-weight: 500;
}

.input-group {
  display: flex;
  align-items: center;
  gap: 8px;
}

.input-group input {
  flex: 1;
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 14px;
}

.input-unit {
  flex: 0 0 30px;
  font-size: 0.85em;
  color: #666;
}

.config-hint {
  margin-top: 12px;
  padding: 10px 14px;
  background: #fff3cd;
  border-left: 3px solid #ffc107;
  border-radius: 4px;
  font-size: 0.85em;
  color: #856404;
}

.config-actions {
  margin-top: 15px;
  text-align: center;
}

.btn-primary {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  border: none;
  padding: 10px 30px;
  border-radius: 6px;
  cursor: pointer;
  font-size: 14px;
  font-weight: 500;
  transition: all 0.3s;
}

.btn-primary:hover:not(:disabled) {
  transform: translateY(-2px);
  box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
}

.btn-primary:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}
</style>
