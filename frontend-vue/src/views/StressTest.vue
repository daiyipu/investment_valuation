<template>
  <div class="stress-test">
    <div class="header">
      <h1>⚠️ 压力测试</h1>
      <p>评估极端情景下的估值变化</p>
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
          <h4 class="config-title">收入冲击测试设置</h4>
          <div class="params-config-grid">
            <div class="config-item">
              <label>收入下降幅度1 (%)</label>
              <div class="input-group">
                <input v-model.number="revenueShockLevels[0]" type="number" step="5" min="5" max="50" />
                <span class="input-unit">%</span>
              </div>
            </div>
            <div class="config-item">
              <label>收入下降幅度2 (%)</label>
              <div class="input-group">
                <input v-model.number="revenueShockLevels[1]" type="number" step="5" min="5" max="50" />
                <span class="input-unit">%</span>
              </div>
            </div>
            <div class="config-item">
              <label>收入下降幅度3 (%)</label>
              <div class="input-group">
                <input v-model.number="revenueShockLevels[2]" type="number" step="5" min="5" max="50" />
                <span class="input-unit">%</span>
              </div>
            </div>
            <div class="config-item">
              <label>收入下降幅度4 (%)</label>
              <div class="input-group">
                <input v-model.number="revenueShockLevels[3]" type="number" step="5" min="5" max="50" />
                <span class="input-unit">%</span>
              </div>
            </div>
          </div>
        </div>

        <div class="config-section">
          <h4 class="config-title">蒙特卡洛模拟设置</h4>
          <div class="config-item">
            <label>模拟次数</label>
            <div class="input-group">
              <input v-model.number="monteCarloIterations" type="number" step="100" min="100" max="10000" />
              <span class="input-unit">次</span>
            </div>
          </div>
          <div class="config-hint">
            💡 模拟次数越多，结果越精确，但计算时间也会增加
          </div>
        </div>

        <div class="config-actions">
          <button @click="runStressTest" class="btn-primary" :disabled="loading">
            {{ loading ? '测试中...' : '重新运行测试' }}
          </button>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">收入冲击测试</div>
      <div class="stress-grid">
        <div v-for="test in revenueShocks" :key="test.scenario_description" class="stress-item">
          <div class="stress-scenario">{{ test.scenario_description }}</div>
          <div class="stress-value">
            {{ formatMoney(test.stressed_value) }}
            <span :class="test.change_pct < 0 ? 'negative' : 'positive'">
              ({{ test.change_pct > 0 ? '+' : '' }}{{ formatPercent(test.change_pct) }})
            </span>
          </div>
        </div>
      </div>
    </div>

    <div v-if="monteCarloData" class="card">
      <div class="card-title">🎲 蒙特卡洛模拟</div>
      <div ref="monteCarloChart" class="chart"></div>
      <div class="monte-carlo-stats">
        <div class="stat-item">
          <span class="stat-label">均值</span>
          <span class="stat-value">{{ formatMoney(monteCarloData.mean) }}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">中值</span>
          <span class="stat-value">{{ formatMoney(monteCarloData.median) }}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">标准差</span>
          <span class="stat-value">{{ formatMoney(monteCarloData.std) }}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">95%置信区间</span>
          <span class="stat-value">
            {{ formatMoney(monteCarloData.percentile_5) }} - {{ formatMoney(monteCarloData.percentile_95) }}
          </span>
        </div>
        <div class="stat-item">
          <span class="stat-label">模拟次数</span>
          <span class="stat-value">{{ monteCarloData.iterations || 1000 }} 次</span>
        </div>
      </div>
    </div>

    <div v-else class="card">
      <div class="card-title">🎲 蒙特卡洛模拟</div>
      <div class="no-data">暂无蒙特卡洛模拟数据</div>
    </div>

    <div class="card">
      <div class="card-title">风险提示</div>
      <div class="risk-summary">
        <div class="risk-item">
          <span class="risk-icon">⚠️</span>
          <span class="risk-text">最大下行风险：{{ maxDownside }}</span>
        </div>
        <div class="risk-item" v-if="revenueShocks.length > 0">
          <span class="risk-icon">📉</span>
          <span class="risk-text">在收入下降30%的情景下，估值下降约{{ getRevenueShockImpact() }}</span>
        </div>
        <div class="risk-item">
          <span class="risk-icon">📊</span>
          <span class="risk-text">建议设置安全边际，当前估值应考虑折扣</span>
        </div>
        <div class="risk-item" v-if="monteCarloData">
          <span class="risk-icon">🎲</span>
          <span class="risk-text">
            蒙特卡洛模拟显示估值分布在 {{ formatMoney(monteCarloData.percentile_5) }} 到 {{ formatMoney(monteCarloData.percentile_95) }}
          </span>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, computed, nextTick } from 'vue'
import * as echarts from 'echarts'
import { stressTestAPI } from '../services/api'

const revenueShocks = ref<any[]>([])
const monteCarloData = ref<any>(null)
const monteCarloChart = ref<HTMLElement>()
const showAdvancedConfig = ref(false)
const loading = ref(false)

const revenueShockLevels = ref([10, 20, 30, 50])  // 默认收入下降幅度
const monteCarloIterations = ref(1000)  // 默认模拟次数

const maxDownside = computed(() => {
  if (revenueShocks.value.length > 0) {
    const minChange = Math.min(...revenueShocks.value.map(t => t.change_pct))
    return (minChange * 100).toFixed(1) + '%'
  }
  return '--'
})

onMounted(async () => {
  const data = sessionStorage.getItem('valuationResults')
  if (data) {
    const parsed = JSON.parse(data)

    if (parsed.stress?.report?.tests?.revenue_shock) {
      revenueShocks.value = parsed.stress.report.tests.revenue_shock
    }

    // 蒙特卡洛数据在 report 级别，不是 tests 级别
    if (parsed.stress?.report?.monte_carlo) {
      monteCarloData.value = parsed.stress.report.monte_carlo
      await nextTick()
      initMonteCarloChart()
    }
  }
})

const initMonteCarloChart = () => {
  if (!monteCarloChart.value || !monteCarloData.value?.distribution) return

  const chart = echarts.init(monteCarloChart.value)
  const distribution = monteCarloData.value.distribution

  // 计算每个区间的中点
  const bins = distribution.map((d: any) => (d.bin_lower + d.bin_upper) / 2 / 10000)
  const counts = distribution.map((d: any) => d.count)

  chart.setOption({
    title: { text: '蒙特卡洛模拟估值分布', left: 'center' },
    tooltip: {
      trigger: 'axis',
      formatter: (params: any) => {
        const idx = params[0].dataIndex
        const bin = distribution[idx]
        return `估值: ${bins[idx].toFixed(2)} 亿元<br/>频次: ${counts[idx]}<br/>区间: ${(bin.bin_lower/10000).toFixed(2)}-${(bin.bin_upper/10000).toFixed(2)} 亿元`
      }
    },
    grid: {
      left: '60px',
      right: '20px',
      bottom: '60px',
      top: '60px'
    },
    xAxis: {
      type: 'category',
      data: bins.map((b: number) => b.toFixed(1)),
      name: '估值（亿元）',
      nameLocation: 'middle',
      nameGap: 30,
      axisLabel: {
        rotate: 45,
        interval: 0
      }
    },
    yAxis: {
      type: 'value',
      name: '频次'
    },
    series: [{
      type: 'bar',
      data: counts,
      itemStyle: { color: '#764ba2' }
    }]
  })

  // 响应式调整
  window.addEventListener('resize', () => {
    const c = echarts.getInstanceByDom(monteCarloChart.value!)
    c?.resize()
  })
}

const getRevenueShockImpact = () => {
  if (revenueShocks.value.length > 0) {
    const shock = revenueShocks.value.find(t => t.scenario_description.includes('-30%'))
    if (shock) {
      return (shock.change_pct * 100).toFixed(1) + '%'
    }
  }
  return '30%'
}

const runStressTest = async () => {
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
      alert('多产品估值模式暂不支持压力测试。\n请使用单产品估值模式进行压力测试。')
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

    // 构建收入冲击数组（转为负数百分比）
    const shocks = revenueShockLevels.value.map(level => -level / 100)

    // 并行调用收入冲击和蒙特卡洛模拟API
    const [revenueShockResponse, monteCarloResponse] = await Promise.all([
      stressTestAPI.revenueShock(companyForAPI, shocks),
      stressTestAPI.monteCarlo(companyForAPI, monteCarloIterations.value)
    ])

    // 更新数据
    revenueShocks.value = revenueShockResponse.data.report?.tests?.revenue_shock || []
    monteCarloData.value = monteCarloResponse.data.report?.tests?.monte_carlo || null

    // 保存到 sessionStorage
    sessionStorage.setItem('valuationResults', JSON.stringify({
      ...parsed,
      stress: {
        report: {
          tests: {
            revenue_shock: revenueShocks.value,
            monte_carlo: monteCarloData.value
          }
        }
      }
    }))

    // 如果有蒙特卡洛数据，重新初始化图表
    if (monteCarloData.value) {
      await nextTick()
      initMonteCarloChart()
    }
  } catch (error) {
    console.error('压力测试失败:', error)
    alert('压力测试失败，请检查参数设置')
  } finally {
    loading.value = false
  }
}

const formatMoney = (value: number) => (value / 10000).toFixed(2) + ' 亿元'
const formatPercent = (value: number) => (value * 100).toFixed(1) + '%'
</script>

<style scoped>
.stress-test {
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

.stress-grid {
  display: grid;
  gap: 15px;
}

.stress-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 15px;
  background: #f8f9fa;
  border-radius: 8px;
}

.stress-scenario {
  color: #555;
  font-weight: 500;
}

.stress-value {
  font-size: 1.2em;
  color: #667eea;
  font-weight: 600;
}

.stress-value .positive { color: #91cc75; }
.stress-value .negative { color: #ee6666; }

.chart {
  height: 350px;
  margin-top: 20px;
}

.monte-carlo-stats {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 15px;
  margin-top: 20px;
}

.stat-item {
  background: #f8f9fa;
  padding: 15px;
  border-radius: 8px;
  text-align: center;
}

.stat-label {
  display: block;
  color: #666;
  font-size: 0.85em;
  margin-bottom: 8px;
}

.stat-value {
  font-size: 1.1em;
  color: #333;
  font-weight: 600;
}

.no-data {
  text-align: center;
  padding: 40px;
  color: #999;
}

.risk-summary {
  display: grid;
  gap: 15px;
}

.risk-item {
  display: flex;
  align-items: center;
  gap: 15px;
  padding: 15px;
  background: #fff5f5;
  border-left: 4px solid #ee6666;
  border-radius: 4px;
}

.risk-icon {
  font-size: 1.5em;
}

.risk-text {
  color: #555;
  font-size: 0.95em;
  flex: 1;
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
  grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
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
