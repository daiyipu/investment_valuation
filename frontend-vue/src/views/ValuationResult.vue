<template>
  <div class="valuation-result">
    <div class="header">
      <h1>📊 估值结果</h1>
      <p>{{ company?.name }} - {{ company?.industry }}</p>
      <button class="btn-back" @click="$router.push('/valuation')">← 返回</button>
    </div>

    <div v-if="!results" class="no-data">
      <p>暂无估值数据</p>
      <button @click="$router.push('/valuation')" class="btn-primary">开始估值</button>
    </div>

    <template v-else>
      <!-- 相对估值结果 -->
      <div v-if="results.relative && Object.keys(results.relative.results || {}).length > 0" class="card">
        <div class="card-title">📈 相对估值</div>
        <div ref="relativeChart" class="chart"></div>
        <div class="methods-list">
          <div v-for="(result, method) in results.relative.results" :key="method" class="method-item">
            <div class="method-header">
              <span class="method-name">{{ getMethodName(method) }}</span>
              <span class="method-value">{{ formatMoney(result.value) }}</span>
            </div>
            <div v-if="result.value_low && result.value_high" class="method-details">
              区间: {{ formatMoney(result.value_low) }} - {{ formatMoney(result.value_high) }}
            </div>
          </div>
        </div>
      </div>

      <!-- 无相对估值数据提示 -->
      <div v-else-if="!results.relative" class="card info-card">
        <div class="card-title">📈 相对估值</div>
        <div class="info-message">
          <p>未添加可比公司，无法进行相对估值</p>
          <p class="hint">相对估值需要可比公司的P/E、P/S、P/B等估值倍数数据</p>
          <button @click="$router.push('/valuation')" class="btn-primary">返回添加可比公司</button>
        </div>
      </div>

      <!-- DCF估值结果 -->
      <div class="card">
        <div class="card-title">💰 DCF绝对估值</div>
        <div class="result-highlight">
          <span class="label">股权价值</span>
          <span class="value">{{ formatMoney(results.dcf?.result?.value) }}</span>
        </div>
        <div class="result-grid">
          <div class="result-item">
            <span class="result-label">WACC</span>
            <span class="result-value">{{ formatPercent(results.dcf?.result?.details?.wacc) }}</span>
          </div>
          <div class="result-item">
            <span class="result-label">预测期现值</span>
            <span class="result-value">{{ formatMoney(results.dcf?.result?.details?.pv_forecasts) }}</span>
          </div>
          <div class="result-item">
            <span class="result-label">终值现值</span>
            <span class="result-value">{{ formatMoney(results.dcf?.result?.details?.pv_terminal) }}</span>
          </div>
          <div class="result-item">
            <span class="result-label">终值占比</span>
            <span class="result-value">{{ getTerminalPercent() }}%</span>
          </div>
        </div>
      </div>

      <!-- 综合估值对比 -->
      <div v-if="hasMultipleValuations" class="card">
        <div class="card-title">📊 综合估值对比</div>
        <div ref="comparisonChart" class="chart"></div>
        <div class="recommendation">
          <div class="rec-label">推荐估值（中位数）</div>
          <div class="rec-value">{{ formatMoney(getRecommendedValue()) }}</div>
          <div class="rec-range">估值区间: {{ getValueRange() }}</div>
        </div>
      </div>

      <!-- 情景分析 -->
      <div class="card">
        <div class="card-title">📈 情景分析</div>
        <div ref="scenarioChart" class="chart"></div>
      </div>

      <!-- 敏感性分析 -->
      <div class="card">
        <div class="card-title">📊 参数敏感性分析</div>
        <div ref="tornadoChart" class="chart"></div>
      </div>

      <!-- 压力测试 -->
      <div class="card">
        <div class="card-title">⚠️ 压力测试结果</div>
        <div class="stress-results">
          <div v-for="(test, idx) in results.stress?.report?.tests?.revenue_shock" :key="idx"
               class="stress-item">
            <span class="stress-label">{{ test.scenario_description }}</span>
            <span class="stress-value">
              {{ formatMoney(test.stressed_value) }}
              <span :class="test.change_pct < 0 ? 'negative' : 'positive'">
                ({{ test.change_pct > 0 ? '+' : '' }}{{ formatPercent(test.change_pct) }})
              </span>
            </span>
          </div>
        </div>
      </div>

      <!-- 蒙特卡洛模拟 -->
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
        </div>
      </div>
    </template>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, nextTick, computed } from 'vue'
import * as echarts from 'echarts'

const results = ref<any>(null)
const company = ref<any>(null)
const relativeChart = ref<HTMLElement>()
const comparisonChart = ref<HTMLElement>()
const scenarioChart = ref<HTMLElement>()
const tornadoChart = ref<HTMLElement>()
const monteCarloChart = ref<HTMLElement>()

const monteCarloData = computed(() => {
  return results.value?.stress?.report?.monte_carlo || null
})

const hasMultipleValuations = computed(() => {
  const hasRelative = results.value?.relative && Object.keys(results.value.relative.results || {}).length > 0
  const hasDCF = results.value?.dcf?.result?.value
  return hasRelative && hasDCF
})

onMounted(async () => {
  const data = sessionStorage.getItem('valuationResults')
  if (data) {
    const parsed = JSON.parse(data)
    results.value = parsed
    company.value = parsed.company

    await nextTick()
    initCharts()
  }
})

const initCharts = () => {
  if (!results.value) return

  // 初始化相对估值图表
  if (relativeChart.value && results.value.relative?.results) {
    const chart = echarts.init(relativeChart.value)
    const methods: string[] = []
    const values: number[] = []

    for (const [method, result] of Object.entries(results.value.relative.results)) {
      methods.push(getMethodName(method))
      values.push(((result as any).value || 0) / 10000)
    }

    if (methods.length > 0) {
      chart.setOption({
        title: { text: '相对估值方法对比', left: 'center' },
        tooltip: { trigger: 'axis', formatter: '{b}: {c} 亿元' },
        xAxis: { type: 'category', data: methods },
        yAxis: { type: 'value', name: '估值（亿元）' },
        series: [{
          type: 'bar',
          data: values,
          itemStyle: { color: '#667eea' }
        }]
      })
    }
  }

  // 初始化综合估值对比图表
  if (comparisonChart.value && hasMultipleValuations.value) {
    const chart = echarts.init(comparisonChart.value)
    const methods: string[] = []
    const values: number[] = []

    // 添加相对估值方法
    if (results.value.relative?.results) {
      for (const [method, result] of Object.entries(results.value.relative.results)) {
        methods.push(getMethodName(method))
        values.push(((result as any).value || 0) / 10000)
      }
    }

    // 添加DCF
    if (results.value.dcf?.result?.value) {
      methods.push('DCF')
      values.push(results.value.dcf.result.value / 10000)
    }

    chart.setOption({
      title: { text: '多方法估值对比', left: 'center' },
      tooltip: { trigger: 'axis', formatter: '{b}: {c} 亿元' },
      xAxis: { type: 'category', data: methods },
      yAxis: { type: 'value', name: '估值（亿元）' },
      series: [{
        type: 'bar',
        data: values,
        itemStyle: {
          color: (params: any) => {
            const colors = ['#5470c6', '#91cc75', '#fac858', '#ee6666', '#73c0de']
            return colors[params.dataIndex % colors.length]
          }
        }
      }]
    })
  }

  // 初始化情景分析图表
  if (scenarioChart.value) {
    const chart = echarts.init(scenarioChart.value)
    const scenarios: string[] = []
    const values: number[] = []

    for (const [name, result] of Object.entries(results.value.scenario?.results || {})) {
      if (name !== 'statistics') {
        const data = result as any
        scenarios.push(name)
        values.push((data.valuation?.value || data.value || 0) / 10000)
      }
    }

    if (scenarios.length > 0) {
      chart.setOption({
        title: { text: '情景分析对比', left: 'center' },
        tooltip: { trigger: 'axis', formatter: '{b}: {c} 亿元' },
        xAxis: { type: 'category', data: scenarios },
        yAxis: { type: 'value', name: '估值（亿元）' },
        series: [{
          type: 'bar',
          data: values,
          itemStyle: {
            color: (params: any) => {
              const colors = ['#91cc75', '#5470c6', '#ee6666']
              return colors[params.dataIndex % colors.length]
            }
          }
        }]
      })
    }
  }

  // 初始化龙卷风图
  if (tornadoChart.value) {
    const chart = echarts.init(tornadoChart.value)
    const sensitivityData = results.value.sensitivity?.results

    if (sensitivityData?.parameters) {
      const params: string[] = []
      const impacts: number[] = []

      for (const [paramName, paramData] of Object.entries(sensitivityData.parameters)) {
        const data = paramData as any
        if (data.valuation_range) {
          params.push(paramName)
          impacts.push((data.valuation_range / 2 / 10000).toFixed(2) as any)
        }
      }

      if (params.length > 0) {
        chart.setOption({
          title: { text: '参数敏感性（估值波动范围）', left: 'center' },
          tooltip: { trigger: 'axis', formatter: '{b}: ±{c} 亿元' },
          xAxis: { type: 'value', name: '估值波动（亿元）' },
          yAxis: { type: 'category', data: params },
          series: [{
            type: 'bar',
            data: impacts,
            itemStyle: { color: '#667eea' }
          }]
        })
      }
    }
  }

  // 初始化蒙特卡洛图表
  if (monteCarloChart.value && monteCarloData.value?.distribution) {
    const chart = echarts.init(monteCarloChart.value)
    const distribution = monteCarloData.value.distribution

    const bins = distribution.map((d: any) => (d.bin_lower + d.bin_upper) / 2 / 10000)
    const counts = distribution.map((d: any) => d.count)

    chart.setOption({
      title: { text: '蒙特卡洛模拟估值分布', left: 'center' },
      tooltip: { trigger: 'axis' },
      xAxis: { type: 'category', data: bins.map(b => b.toFixed(1)), name: '估值（亿元）' },
      yAxis: { type: 'value', name: '频次' },
      series: [{
        type: 'bar',
        data: counts,
        itemStyle: { color: '#764ba2' }
      }]
    })
  }

  // 响应式调整
  window.addEventListener('resize', () => {
    if (relativeChart.value) {
      const c = echarts.getInstanceByDom(relativeChart.value)
      c?.resize()
    }
    if (comparisonChart.value) {
      const c = echarts.getInstanceByDom(comparisonChart.value)
      c?.resize()
    }
    if (scenarioChart.value) {
      const c = echarts.getInstanceByDom(scenarioChart.value)
      c?.resize()
    }
    if (tornadoChart.value) {
      const c = echarts.getInstanceByDom(tornadoChart.value)
      c?.resize()
    }
    if (monteCarloChart.value) {
      const c = echarts.getInstanceByDom(monteCarloChart.value)
      c?.resize()
    }
  })
}

const getMethodName = (method: string) => {
  const names: Record<string, string> = {
    'P/E法': '市盈率法 (P/E)',
    'P/S法': '市销率法 (P/S)',
    'P/B法': '市净率法 (P/B)',
    'EV/EBITDA法': 'EV/EBITDA倍数法',
    '综合': '综合估值'
  }
  return names[method] || method
}

const getRecommendedValue = () => {
  const values: number[] = []

  if (results.value?.relative?.results) {
    for (const result of Object.values(results.value.relative.results)) {
      values.push((result as any).value || 0)
    }
  }

  if (results.value?.dcf?.result?.value) {
    values.push(results.value.dcf.result.value)
  }

  if (values.length === 0) return 0

  // 返回中位数
  values.sort((a, b) => a - b)
  return values[Math.floor(values.length / 2)]
}

const getValueRange = () => {
  const values: number[] = []

  if (results.value?.relative?.results) {
    for (const result of Object.values(results.value.relative.results)) {
      values.push((result as any).value || 0)
    }
  }

  if (results.value?.dcf?.result?.value) {
    values.push(results.value.dcf.result.value)
  }

  if (values.length === 0) return '--'

  const min = Math.min(...values) * 0.9 / 10000
  const max = Math.max(...values) * 1.1 / 10000
  return `${min.toFixed(2)} - ${max.toFixed(2)} 亿元`
}

const getTerminalPercent = () => {
  const pvTerminal = results.value?.dcf?.result?.details?.pv_terminal || 0
  const total = results.value?.dcf?.result?.value || 1
  return ((pvTerminal / total) * 100).toFixed(1)
}

const formatMoney = (value: number | undefined) => {
  if (!value) return '--'
  return (value / 10000).toFixed(2) + ' 亿元'
}

const formatPercent = (value: number | undefined) => {
  if (!value) return '--'
  return (value * 100).toFixed(2) + '%'
}
</script>

<style scoped>
.valuation-result {
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
  position: relative;
}

.header h1 {
  font-size: 2em;
  margin-bottom: 10px;
}

.btn-back {
  position: absolute;
  left: 20px;
  top: 50%;
  transform: translateY(-50%);
  background: rgba(255, 255, 255, 0.2);
  border: 1px solid rgba(255, 255, 255, 0.4);
  color: white;
  padding: 8px 16px;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.3s;
}

.btn-back:hover {
  background: rgba(255, 255, 255, 0.3);
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

.card-title {
  font-size: 1.3em;
  color: #333;
  margin-bottom: 20px;
  padding-bottom: 10px;
  border-bottom: 2px solid #667eea;
}

.chart {
  height: 350px;
  margin-top: 20px;
}

.methods-list {
  display: flex;
  flex-direction: column;
  gap: 12px;
  margin-top: 20px;
}

.method-item {
  background: #f8f9fa;
  padding: 15px;
  border-radius: 8px;
}

.method-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 8px;
}

.method-name {
  font-weight: bold;
  color: #333;
}

.method-value {
  font-size: 1.3em;
  color: #667eea;
  font-weight: 600;
}

.method-details {
  font-size: 0.9em;
  color: #666;
}

.result-highlight {
  text-align: center;
  padding: 20px;
  background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%);
  border-radius: 8px;
  margin-bottom: 20px;
}

.result-highlight .label {
  display: block;
  color: #666;
  font-size: 0.9em;
  margin-bottom: 10px;
}

.result-highlight .value {
  font-size: 2em;
  color: #667eea;
  font-weight: bold;
}

.result-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 15px;
}

.result-item {
  background: #f8f9fa;
  padding: 15px;
  border-radius: 8px;
  text-align: center;
}

.result-label {
  display: block;
  color: #666;
  font-size: 0.85em;
  margin-bottom: 8px;
}

.result-value {
  font-size: 1.2em;
  color: #333;
  font-weight: 500;
}

.recommendation {
  margin-top: 20px;
  padding: 20px;
  background: linear-gradient(135deg, #f0f7ff 0%, #e8f4ff 100%);
  border-radius: 8px;
  text-align: center;
}

.rec-label {
  color: #666;
  font-size: 0.9em;
  margin-bottom: 10px;
}

.rec-value {
  font-size: 2em;
  color: #667eea;
  font-weight: bold;
  margin-bottom: 10px;
}

.rec-range {
  color: #555;
  font-size: 1em;
}

.stress-results {
  display: grid;
  gap: 10px;
}

.stress-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 15px;
  background: #f8f9fa;
  border-radius: 8px;
}

.stress-label {
  color: #555;
  font-size: 0.95em;
}

.stress-value {
  font-size: 1.1em;
  color: #667eea;
  font-weight: 500;
}

.stress-value .positive { color: #91cc75; }
.stress-value .negative { color: #ee6666; }

.monte-carlo-stats {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
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
  font-size: 0.9em;
  margin-bottom: 8px;
}

.stat-value {
  font-size: 1.1em;
  color: #333;
  font-weight: 600;
}

.btn-primary {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  border: none;
  padding: 12px 30px;
  border-radius: 6px;
  font-size: 16px;
  cursor: pointer;
}

.info-card {
  background: white;
  padding: 25px;
  border-radius: 12px;
  margin-bottom: 20px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
  border-left: 4px solid #ffa500;
}

.info-message {
  text-align: center;
  padding: 20px;
}

.info-message p {
  margin: 10px 0;
  color: #666;
}

.info-message .hint {
  color: #999;
  font-size: 0.9em;
}
</style>
