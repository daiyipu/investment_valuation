<template>
  <div class="report">
    <div class="header">
      <h1>📄 综合估值报告</h1>
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
          <div class="method-item">
            <div class="method-header">
              <span class="method-name">DCF现金流折现</span>
              <span class="method-value">{{ formatMoney(results.dcf?.result?.value) }}</span>
            </div>
            <div class="method-details">
              WACC: {{ formatPercent(results.dcf?.result?.details?.wacc) }} |
              终值占比: {{ getTerminalValuePercent() }}%
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
          <p class="risk-text">
            在收入下降30%的极端情景下，估值下降约
            <strong>{{ getStressImpact() }}</strong>
          </p>
        </div>

        <div class="risk-section">
          <h3>敏感性分析</h3>
          <p class="risk-text">
            <strong>{{ getMostSensitive() }}</strong>是对估值影响最大的参数
          </p>
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
        <button @click="exportReport" class="btn-primary">📥 导出报告</button>
        <button @click="$router.push('/valuation')" class="btn-secondary">重新估值</button>
      </div>
    </template>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, computed } from 'vue'

const results = ref<any>(null)
const company = ref<any>(null)

onMounted(() => {
  const data = sessionStorage.getItem('valuationResults')
  if (data) {
    const parsed = JSON.parse(data)
    results.value = parsed
    company.value = parsed.company
  }
})

const getRecommendedValue = () => {
  return results.value?.dcf?.result?.value || 0
}

const getValueRange = () => {
  const value = getRecommendedValue()
  const low = (value * 0.9 / 10000).toFixed(2)
  const high = (value * 1.1 / 10000).toFixed(2)
  return `${low} - ${high} 亿元`
}

const getTerminalValuePercent = () => {
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

const getStressImpact = () => {
  const shocks = results.value?.stress?.report?.tests?.revenue_shock
  if (shocks && shocks.length > 0) {
    const impact = shocks[0].change_pct
    return (impact * 100).toFixed(1) + '%'
  }
  return '--'
}

const getMaxDownside = () => {
  const shocks = results.value?.stress?.report?.tests?.revenue_shock
  if (shocks && shocks.length > 0) {
    const minChange = Math.min(...shocks.map((t: any) => t.change_pct))
    return (minChange * 100).toFixed(1) + '%'
  }
  return '--'
}

const getMostSensitive = () => {
  const params = results.value?.sensitivity?.results?.parameters
  if (params) {
    const entries = Object.entries(params).sort((a: any, b: any) =>
      b[1].valuation_range - a[1].valuation_range)
    return entries[0]?.[0] || '--'
  }
  return '--'
}

const getRecommendationLevel = () => {
  return 'hold'
}

const getRecommendationText = () => {
  return '中性持有'
}

const getSafetyMargin = () => '20%'

const formatMoney = (value: number) => (value / 10000).toFixed(2) + ' 亿元'
const formatPercent = (value: number) => (value * 100).toFixed(2) + '%'

const exportReport = () => {
  alert('报告导出功能开发中...\n\n可复制页面内容或使用浏览器打印功能保存为PDF')
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
  padding: 60px;
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
  align-items: center;
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

.risk-text {
  color: #555;
  line-height: 1.6;
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
