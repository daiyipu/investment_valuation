<template>
  <div class="scenario-analysis">
    <div class="header">
      <h1>📈 情景分析</h1>
      <p>基准/乐观/悲观情景估值对比</p>
    </div>

    <!-- 参数配置 -->
    <div class="card config-card">
      <div class="card-title">
        ⚙️ 参数配置
        <button @click="showConfig = !showConfig" class="btn-toggle">
          {{ showConfig ? '收起 ▲' : '展开 ▼' }}
        </button>
      </div>
      <div v-show="showConfig" class="config-content">
        <div class="scenario-config-list">
          <div class="scenario-item">
            <div class="scenario-label">乐观情景</div>
            <div class="scenario-inputs">
              <div class="input-group">
                <label>收入增长调整</label>
                <input v-model.number="params.bull.revenue_growth_adj" type="number" step="5" />
                <span class="input-unit">%</span>
              </div>
              <div class="input-group">
                <label>利润率调整</label>
                <input v-model.number="params.bull.margin_adj" type="number" step="1" />
                <span class="input-unit">%</span>
              </div>
              <div class="input-group">
                <label>WACC调整</label>
                <input v-model.number="params.bull.wacc_adj" type="number" step="0.5" />
                <span class="input-unit">%</span>
              </div>
            </div>
          </div>

          <div class="scenario-item">
            <div class="scenario-label">悲观情景</div>
            <div class="scenario-inputs">
              <div class="input-group">
                <label>收入增长调整</label>
                <input v-model.number="params.bear.revenue_growth_adj" type="number" step="5" />
                <span class="input-unit">%</span>
              </div>
              <div class="input-group">
                <label>利润率调整</label>
                <input v-model.number="params.bear.margin_adj" type="number" step="1" />
                <span class="input-unit">%</span>
              </div>
              <div class="input-group">
                <label>WACC调整</label>
                <input v-model.number="params.bear.wacc_adj" type="number" step="0.5" />
                <span class="input-unit">%</span>
              </div>
            </div>
          </div>
        </div>
        <div class="config-hint">
          💡 参数调整说明：正数表示增加，负数表示减少。例如：收入增长调整 +20% 表示在基准增长率基础上增加20个百分点
        </div>
        <div class="config-actions">
          <button @click="runAnalysis" class="btn-primary" :disabled="loading">
            {{ loading ? '分析中...' : '重新运行分析' }}
          </button>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">情景对比</div>
      <div ref="scenarioChart" class="chart"></div>
    </div>

    <div class="card">
      <div class="card-title">情景详情</div>
      <div class="scenario-table-container">
        <table class="scenario-table">
          <thead>
            <tr>
              <th>情景</th>
              <th>估值结果</th>
              <th>收入增长调整</th>
              <th>利润率调整</th>
              <th>WACC调整</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(scenario, name) in scenarios" :key="name" v-show="name !== 'statistics'" :class="getScenarioClass(name)">
              <td class="scenario-name-cell">
                <span class="scenario-badge" :class="getScenarioClass(name)">{{ name }}</span>
              </td>
              <td class="scenario-value-cell">{{ formatMoney(scenario.valuation?.value || scenario.value) }}</td>
              <td class="scenario-param-cell">
                <span v-if="scenario.scenario && scenario.scenario.revenue_growth_adj !== undefined"
                      :class="getParamClass(scenario.scenario.revenue_growth_adj)">
                  {{ formatPercent(scenario.scenario.revenue_growth_adj) }}
                </span>
                <span v-else>--</span>
              </td>
              <td class="scenario-param-cell">
                <span v-if="scenario.scenario && scenario.scenario.margin_adj !== undefined"
                      :class="getParamClass(scenario.scenario.margin_adj)">
                  {{ formatPercent(scenario.scenario.margin_adj) }}
                </span>
                <span v-else>--</span>
              </td>
              <td class="scenario-param-cell">
                <span v-if="scenario.scenario && scenario.scenario.wacc_adj !== undefined"
                      :class="getParamClass(scenario.scenario.wacc_adj, true)">
                  {{ formatPercent(scenario.scenario.wacc_adj) }}
                </span>
                <span v-else>--</span>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, nextTick } from 'vue'
import * as echarts from 'echarts'
import { scenarioAPI } from '../services/api'

const scenarios = ref<any>({})
const scenarioChart = ref<HTMLElement>()
const chartInstance = ref<echarts.ECharts | null>(null)
const showConfig = ref(false)
const loading = ref(false)

const params = ref({
  bull: {
    revenue_growth_adj: 20,
    margin_adj: 5,
    wacc_adj: -1
  },
  bear: {
    revenue_growth_adj: -20,
    margin_adj: -5,
    wacc_adj: 2
  }
})

onMounted(async () => {
  const data = sessionStorage.getItem('valuationResults')
  if (data) {
    const parsed = JSON.parse(data)

    // 支持两种数据格式：完整估值结果格式和历史记录格式
    if (parsed.company) {
      // 完整估值结果格式
      if (!parsed.scenario) {
        try {
          const response = await scenarioAPI.analyze(parsed.company)
          scenarios.value = response.data.results
          sessionStorage.setItem('valuationResults', JSON.stringify({
            ...parsed,
            scenario: response.data
          }))
        } catch (error) {
          console.error('获取情景分析失败:', error)
        }
      } else {
        scenarios.value = parsed.scenario.results
      }
    } else if (parsed.results) {
      // 历史记录格式：直接使用保存的情景数据，不重新调用 API
      scenarios.value = parsed.results
    }

    initChart()
  }
})

const initChart = () => {
  if (!scenarioChart.value) return

  // 销毁旧实例
  if (chartInstance.value) {
    chartInstance.value.dispose()
  }

  const chart = echarts.init(scenarioChart.value)
  chartInstance.value = chart
  const names: string[] = []
  const values: number[] = []

  for (const [name, data] of Object.entries(scenarios.value)) {
    if (name !== 'statistics') {
      names.push(name)
      const val = (data as any).valuation?.value || (data as any).value || 0
      // 如果值大于100万，说明单位是万元，需要转换为亿元
      values.push(val > 1000000 ? val / 10000 : val / 10000)
    }
  }

  chart.setOption({
    title: { text: '情景估值对比', left: 'center' },
    tooltip: { trigger: 'axis', formatter: '{b}: {c} 亿元' },
    xAxis: { type: 'category', data: names },
    yAxis: { type: 'value', name: '估值（亿元）' },
    series: [{
      type: 'bar',
      data: values,
      itemStyle: {
        color: (params: any) => {
          const colors = { '基准': '#5470c6', '乐观': '#91cc75', '悲观': '#ee6666' }
          return colors[params.name as keyof typeof colors] || '#5470c6'
        }
      }
    }]
  })
}

const getScenarioClass = (name: string) => {
  if (name === '乐观') return 'bull'
  if (name === '悲观') return 'bear'
  return 'base'
}

const getParamClass = (value: number, isInvert: boolean = false) => {
  const numValue = typeof value === 'number' ? value : parseFloat(value)
  if (isInvert) {
    // WACC: 负值是好事（降低成本），正值是坏事
    return numValue < 0 ? 'param-positive' : 'param-negative'
  } else {
    // 其他参数：正值是好事，负值是坏事
    return numValue > 0 ? 'param-positive' : 'param-negative'
  }
}

const formatMoney = (value: number) => ((value || 0) / 10000).toFixed(2) + ' 亿元'
const formatPercent = (value: number | string) => {
  const numValue = typeof value === 'number' ? value : parseFloat(value)
  return (numValue * 100).toFixed(1) + '%'
}

const runAnalysis = async () => {
  loading.value = true
  try {
    const data = sessionStorage.getItem('valuationResults')
    if (!data) {
      alert('请先进行估值分析')
      return
    }

    const parsed = JSON.parse(data)
    const company = parsed.company

    // 构建情景参数数组（基准情景参数固定为0，乐观和悲观可配置）
    const scenarioParams = [
      {
        name: '基准情景',
        revenue_growth_adj: 0,
        margin_adj: 0,
        wacc_adj: 0
      },
      {
        name: '乐观情景',
        revenue_growth_adj: params.value.bull.revenue_growth_adj / 100,
        margin_adj: params.value.bull.margin_adj / 100,
        wacc_adj: params.value.bull.wacc_adj / 100
      },
      {
        name: '悲观情景',
        revenue_growth_adj: params.value.bear.revenue_growth_adj / 100,
        margin_adj: params.value.bear.margin_adj / 100,
        wacc_adj: params.value.bear.wacc_adj / 100
      }
    ]

    const response = await scenarioAPI.analyze(company, scenarioParams)
    scenarios.value = response.data.results

    // 保存到 sessionStorage
    sessionStorage.setItem('valuationResults', JSON.stringify({
      ...parsed,
      scenario: response.data
    }))

    // 重新初始化图表
    await nextTick()
    initChart()
  } catch (error) {
    console.error('情景分析失败:', error)
    alert('情景分析失败，请检查参数设置')
  } finally {
    loading.value = false
  }
}

</script>

<style scoped>
.scenario-analysis {
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

.config-card {
  /* 高度自适应，宽度保持固定 */
}

.card-title {
  font-size: 1.3em;
  color: #333;
  margin-bottom: 20px;
  padding-bottom: 10px;
  border-bottom: 2px solid #667eea;
}

.chart {
  height: 400px;
}

/* 情景表格样式 */
.scenario-table-container {
  overflow-x: auto;
}

.scenario-table {
  width: 100%;
  border-collapse: collapse;
  background: white;
}

.scenario-table thead {
  background: #f8f9fa;
}

.scenario-table th {
  padding: 12px 15px;
  text-align: center;
  font-weight: 600;
  color: #555;
  border-bottom: 2px solid #e0e0e0;
}

.scenario-table td {
  padding: 15px;
  text-align: center;
  border-bottom: 1px solid #f0f0f0;
}

.scenario-table tbody tr {
  transition: background 0.2s;
}

.scenario-table tbody tr:hover {
  background: #f8f9fa;
}

.scenario-table tbody tr.bull {
  background: #e8ffe8;
}

.scenario-table tbody tr.bear {
  background: #ffe8e8;
}

.scenario-table tbody tr.base {
  background: #e8f0ff;
}

.scenario-name-cell {
  font-weight: 600;
}

.scenario-badge {
  display: inline-block;
  padding: 6px 16px;
  border-radius: 20px;
  font-weight: 600;
}

.scenario-badge.base {
  background: #5470c6;
  color: white;
}

.scenario-badge.bull {
  background: #91cc75;
  color: white;
}

.scenario-badge.bear {
  background: #ee6666;
  color: white;
}

.scenario-value-cell {
  font-size: 1.2em;
  font-weight: bold;
  color: #667eea;
}

.scenario-param-cell span {
  display: inline-block;
  padding: 4px 12px;
  border-radius: 4px;
  font-weight: 500;
}

.param-positive {
  color: #27ae60;
  background: #e8f8e8;
}

.param-negative {
  color: #e74c3c;
  background: #fde8e8;
}

.scenario-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
  gap: 20px;
}

.scenario-card {
  padding: 20px;
  border-radius: 8px;
  text-align: center;
}

.scenario-card.base {
  background: #e8f0ff;
  border: 2px solid #5470c6;
}

.scenario-card.bull {
  background: #e8ffe8;
  border: 2px solid #91cc75;
}

.scenario-card.bear {
  background: #ffe8e8;
  border: 2px solid #ee6666;
}

.scenario-header {
  font-size: 1.2em;
  font-weight: bold;
  margin-bottom: 15px;
  color: #333;
}

.scenario-value {
  font-size: 1.8em;
  font-weight: bold;
  color: #667eea;
  margin-bottom: 15px;
}

.scenario-params {
  font-size: 0.9em;
  color: #666;
  line-height: 1.6;
}

/* 参数配置样式 */
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

.config-content {
  margin-top: 20px;
  min-height: 260px;
  max-width: 600px;
  margin-left: auto;
  margin-right: auto;
  overflow: hidden;
}

.scenario-config-list {
  display: flex;
  flex-direction: row;
  gap: 30px;
  justify-content: center;
}

.scenario-item {
  background: #f8f9fa;
  padding: 12px;
  padding-right: 7px;
  border-radius: 6px;
  border-left: 4px solid #667eea;
  width: 235px;
  flex-shrink: 0;
}

.scenario-label {
  font-weight: 600;
  color: #333;
  margin-bottom: 12px;
  font-size: 0.95em;
}

.scenario-inputs {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.input-group {
  display: flex;
  align-items: center;
  gap: 8px;
}

.input-group label {
  flex: 0 0 80px;
  font-size: 0.85em;
  color: #555;
  white-space: nowrap;
}

.input-group input {
  flex: 1;
  min-width: 30px;
  max-width: 60px;
  padding: 2px 3px;
  border: 1px solid #ddd;
  border-radius: 3px;
  font-size: 12px;
  text-align: center;
  box-sizing: border-box;
}

.input-unit {
  flex: 0 0 20px;
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
  line-height: 1.5;
  word-wrap: break-word;
  overflow-wrap: break-word;
  max-width: 500px;
  margin-left: auto;
  margin-right: auto;
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
