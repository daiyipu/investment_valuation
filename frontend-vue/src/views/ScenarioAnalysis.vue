<template>
  <div class="scenario-analysis">
    <div class="header">
      <h1>📈 情景分析</h1>
      <p>基准/乐观/悲观情景估值对比</p>
    </div>

    <div class="card">
      <div class="card-title">情景对比</div>
      <div ref="scenarioChart" class="chart"></div>
    </div>

    <div class="card">
      <div class="card-title">情景详情</div>
      <div class="scenario-grid">
        <div v-for="(scenario, name) in scenarios" :key="name" class="scenario-card"
             :class="getScenarioClass(name)">
          <div class="scenario-header">{{ name }}</div>
          <div class="scenario-value">{{ formatMoney(scenario.valuation?.value || scenario.value) }}</div>
          <div v-if="scenario.scenario" class="scenario-params">
            <div v-if="scenario.scenario.revenue_growth_adj !== undefined">
              收入增长调整: {{ formatPercent(scenario.scenario.revenue_growth_adj) }}
            </div>
            <div v-if="scenario.scenario.margin_adj !== undefined">
              利润率调整: {{ formatPercent(scenario.scenario.margin_adj) }}
            </div>
            <div v-if="scenario.scenario.wacc_adj !== undefined">
              WACC调整: {{ formatPercent(scenario.scenario.wacc_adj) }}
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from 'vue'
import * as echarts from 'echarts'
import { scenarioAPI } from '../services/api'

const scenarios = ref<any>({})
const scenarioChart = ref<HTMLElement>()

onMounted(async () => {
  const data = sessionStorage.getItem('valuationResults')
  if (data) {
    const parsed = JSON.parse(data)
    const company = parsed.company

    // 如果没有情景数据，重新获取
    if (!parsed.scenario) {
      try {
        const response = await scenarioAPI.analyze(company)
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

    initChart()
  }
})

const initChart = () => {
  if (!scenarioChart.value) return

  const chart = echarts.init(scenarioChart.value)
  const names: string[] = []
  const values: number[] = []

  for (const [name, data] of Object.entries(scenarios.value)) {
    if (name !== 'statistics') {
      names.push(name)
      values.push(((data as any).valuation?.value || (data as any).value || 0) / 10000)
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

const formatMoney = (value: number) => (value / 10000).toFixed(2) + ' 亿元'
const formatPercent = (value: number) => (value * 100).toFixed(1) + '%'
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
</style>
