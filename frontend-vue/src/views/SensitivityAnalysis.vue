<template>
  <div class="sensitivity-analysis">
    <div class="header">
      <h1>📊 敏感性分析</h1>
      <p>评估各参数对估值的影响程度</p>
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

const sensitivityParams = ref<Record<string, any>>({})
const tornadoChart = ref<HTMLElement>()

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

const initTornadoChart = () => {
  if (!tornadoChart.value) return

  const chart = echarts.init(tornadoChart.value)
  const params: string[] = []
  const impacts: number[] = []

  for (const [paramName, paramData] of Object.entries(sensitivityParams.value)) {
    if (paramData.valuation_range) {
      params.push(paramName)
      impacts.push((paramData.valuation_range / 2 / 10000).toFixed(2) as any)
    }
  }

  // 按影响程度排序
  const sorted = params.map((p, i) => ({ name: p, impact: parseFloat(impacts[i] as string) }))
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
</style>
