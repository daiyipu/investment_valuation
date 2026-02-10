<template>
  <div class="stress-test">
    <div class="header">
      <h1>⚠️ 压力测试</h1>
      <p>评估极端情景下的估值变化</p>
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

const revenueShocks = ref<any[]>([])
const monteCarloData = ref<any>(null)
const monteCarloChart = ref<HTMLElement>()

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

    if (parsed.stress?.report?.tests?.monte_carlo) {
      monteCarloData.value = parsed.stress.report.tests.monte_carlo
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
      data: bins.map(b => b.toFixed(1)),
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
</style>
