<template>
  <div class="valuation-result">
    <div class="header">
      <h1>📊 估值结果</h1>
      <p>{{ company?.name }} - {{ company?.industry }}</p>
      <button class="btn-back" @click="$router.push('/valuation')">← 返回</button>
    </div>

    <div v-if="!results" class="no-data">
      <p>⚠️ 暂无估值数据</p>
      <p class="hint">可能的原因：</p>
      <ul class="error-list">
        <li>页面直接访问（请先填写公司数据并开始估值）</li>
        <li>浏览器缓存或sessionStorage被清空</li>
        <li>数据保存失败</li>
      </ul>
      <button @click="$router.push('/valuation')" class="btn-primary">开始估值</button>
    </div>

    <template v-else>
      <!-- 相对估值结果（多产品和单产品都显示） -->
      <div v-if="results.relative && Object.keys(results.relative.results || {}).length > 0" class="card">
        <div class="card-title">📈 相对估值</div>
        <div ref="relativeChart" class="chart"></div>
        <div class="methods-list">
          <div v-for="(result, method) in results.relative.results" :key="String(method)" class="method-item">
            <div class="method-header">
              <span class="method-name">{{ getMethodName(String(method)) }}</span>
              <span class="method-value">{{ formatMoney(result.value) }}</span>
            </div>
            <div v-if="result.value_low && result.value_high" class="method-details">
              区间: {{ formatMoney(result.value_low) }} - {{ formatMoney(result.value_high) }}
            </div>
          </div>
        </div>
      </div>

      <!-- 多产品估值结果 -->
      <div v-if="isMultiProduct && multiProductData" class="multi-product-section">
        <!-- 整体估值概览 -->
        <div class="card">
          <div class="card-title">🏢 多产品估值 - 整体估值</div>
          <div class="result-highlight">
            <span class="label">企业价值</span>
            <span class="value">{{ formatMoney(correctedTotalEnterpriseValue) }}</span>
          </div>
          <div class="result-grid">
            <div class="result-item">
              <span class="result-label">股权价值</span>
              <span class="result-value">{{ formatMoney(correctedTotalEquityValue) }}</span>
            </div>
            <div class="result-item">
              <span class="result-label">WACC</span>
              <span class="result-value">{{ formatPercent(multiProductData.wacc) }}</span>
            </div>
            <div class="result-item">
              <span class="result-label">总营业收入</span>
              <span class="result-value">{{ formatMoney(multiProductData.total_revenue) }}</span>
            </div>
            <div class="result-item">
              <span class="result-label">产品数量</span>
              <span class="result-value">{{ multiProductData.product_results?.length || 0 }} 个</span>
            </div>
          </div>
          <!-- 显示修正说明 -->
          <div style="margin-top: 12px; padding: 8px 12px; background: #e8f4ff; border-left: 3px solid #667eea; border-radius: 4px; font-size: 0.85em; color: #555;">
            💡 企业价值 = 所有产品的企业价值之和 ({{ formatMoney(correctedTotalEnterpriseValue) }})
          </div>
        </div>

        <!-- 产品价值贡献 -->
        <div class="card">
          <div class="card-title">📊 产品价值贡献分析</div>
          <div ref="productContributionChart" class="chart"></div>
          <div class="product-contribution-list">
            <div v-for="contribution in multiProductData.product_contribution" :key="contribution.product"
                 class="contribution-item">
              <div class="contribution-header">
                <span class="contribution-name">{{ contribution.product }}</span>
                <span class="contribution-percent">{{ contribution.contribution_pct.toFixed(1) }}%</span>
              </div>
              <div class="contribution-bar">
                <div class="contribution-fill" :style="{ width: contribution.contribution_pct + '%' }"></div>
              </div>
            </div>
          </div>
        </div>

        <!-- 分产品估值明细 -->
        <div class="card">
          <div class="card-title">📋 分产品估值明细</div>
          <div class="product-table-container">
            <table class="product-table">
              <thead>
                <tr>
                  <th>产品名称</th>
                  <th>收入占比</th>
                  <th>当前收入</th>
                  <th>预测期现值</th>
                  <th>终值现值</th>
                  <th>企业价值</th>
                  <th>终值占比</th>
                  <th>收入CAGR</th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="product in multiProductData.product_results" :key="product.product_name"
                    class="product-row">
                  <td class="product-name">{{ product.product_name }}</td>
                  <td>{{ (product.revenue_weight * 100).toFixed(1) }}%</td>
                  <td>{{ formatMoney(product.current_revenue) }}</td>
                  <td>{{ formatMoney(product.pv_forecasts) }}</td>
                  <td>{{ formatMoney(product.pv_terminal) }}</td>
                  <td class="enterprise-value">{{ formatMoney(product.enterprise_value) }}</td>
                  <td>{{ (product.pv_terminal / product.enterprise_value * 100).toFixed(1) }}%</td>
                  <td>{{ (product.revenue_cagr * 100).toFixed(1) }}%</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div class="info-note">
            💡 <strong>企业价值构成</strong>：预测期现值（5年现金流折现）+ 终值现值（永续增长价值折现）
          </div>
        </div>

        <!-- 综合估值对比（多产品模式） -->
        <div v-if="results.relative && Object.keys(results.relative.results || {}).length > 0" class="card">
          <div class="card-title">📊 综合估值对比</div>
          <div ref="comparisonChart" class="chart"></div>
          <div class="recommendation">
            <div class="rec-label">推荐估值（中位数）</div>
            <div class="rec-value">{{ formatMoney(getRecommendedValueForMultiProduct()) }}</div>
            <div class="rec-range">估值区间: {{ getValueRangeForMultiProduct() }}</div>
          </div>
        </div>

        <!-- 合并现金流预测 -->
        <div class="card">
          <div class="card-title">💰 合并现金流预测</div>
          <div ref="consolidatedCashFlowChart" class="chart"></div>
        </div>

        <!-- 分产品现金流详情 -->
        <div class="card">
          <div class="card-title">📈 分产品现金流预测</div>
          <div v-for="product in multiProductData.product_results" :key="product.product_name"
               class="product-cashflow-section">
            <h4 class="product-section-title">{{ product.product_name }}</h4>
            <div class="product-cashflow-grid">
              <div v-for="forecast in product.fcf_forecasts" :key="forecast.year" class="forecast-item">
                <span class="forecast-year">第{{ forecast.year }}年</span>
                <span class="forecast-revenue">收入: {{ formatMoney(forecast.revenue) }}</span>
                <span class="forecast-fcf">FCF: {{ formatMoney(forecast.fcf) }}</span>
                <span class="forecast-growth">增长: {{ (forecast.growth_rate * 100).toFixed(1) }}%</span>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- 无相对估值数据提示（多产品和单产品都显示） -->
      <div v-if="!results.relative" class="card info-card">
        <div class="card-title">📈 相对估值</div>
        <div class="info-message">
          <!-- 如果有错误信息，显示错误详情 -->
          <div v-if="results.relativeError" style="background: #fee; border-left: 4px solid #f66; padding: 12px; margin-bottom: 12px;">
            <p style="color: #c33; margin: 0 0 8px 0; font-weight: 500;">⚠️ 相对估值获取失败</p>
            <p style="color: #666; margin: 4px 0; font-size: 0.9em;">错误: {{ results.relativeError.message }}</p>
            <div v-if="results.relativeError.response" style="margin-top: 8px;">
              <details style="cursor: pointer; color: #666;">
                <summary>查看API响应详情</summary>
                <pre style="background: #fff; padding: 8px; margin-top: 8px; font-size: 0.85em; overflow-x: auto;">{{ JSON.stringify(results.relativeError.response, null, 2) }}</pre>
              </details>
            </div>
          </div>

          <!-- 如果有可比公司但无结果 -->
          <div v-else-if="results.hasComparables" style="background: #ffeaa7; border-left: 4px solid #fdcb6e; padding: 12px; margin-bottom: 12px;">
            <p style="color: #2d3436; margin: 0 0 8px 0; font-weight: 500;">⚠️ 已添加可比公司但相对估值未成功</p>
            <p style="color: #636e72; margin: 4px 0; font-size: 0.9em;">
              可比公司数量: {{ results.comparables?.length || 0 }}
            </p>
            <p v-if="results.noComparablesReason" style="color: #636e72; margin: 4px 0; font-size: 0.9em;">
              跳过原因: {{ results.noComparablesReason }}
            </p>
            <p style="color: #636e72; margin: 8px 0 0 0; font-size: 0.9em;">
              💡 请检查浏览器控制台查看详细错误日志
            </p>
          </div>

          <!-- 如果没有可比公司 -->
          <div v-else>
            <p>未添加可比公司，无法进行相对估值</p>
            <p class="hint">相对估值需要可比公司的P/E、P/S、P/B等估值倍数数据</p>
            <p class="hint" style="color: #667eea; font-weight: 500;">💡 提示：请返回估值页面，在"单产品估值"模式下添加可比公司，然后重新估值</p>
          </div>
        </div>
      </div>

      <!-- 单产品估值结果（原有逻辑） -->
      <template v-if="!isMultiProduct">
      <!-- DCF估值结果 -->
      <div class="card">
        <div class="card-title">💰 DCF绝对估值 - 企业价值分解</div>
        <div class="result-highlight">
          <span class="label">股权价值</span>
          <span class="value">{{ formatMoney(results.dcf?.result?.value) }}</span>
        </div>
        <div class="result-grid">
          <div class="result-item">
            <span class="result-label">企业价值</span>
            <span class="result-value">{{ formatMoney((results.dcf?.result?.value || 0) + (results.company?.total_debt || 0) - (results.company?.cash_and_equivalents || 0)) }}</span>
          </div>
          <div class="result-item">
            <span class="result-label">WACC</span>
            <span class="result-value">{{ formatPercent(results.dcf?.result?.details?.wacc) }}</span>
          </div>
          <div class="result-item">
            <span class="result-label">当前收入</span>
            <span class="result-value">{{ formatMoney(results.company?.revenue) }}</span>
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
        <div class="info-note">
          💡 <strong>企业价值构成</strong>：预测期现值（5年现金流折现）+ 终值现值（永续增长价值折现）
        </div>
      </div>

      <!-- 价值构成分析（单产品模式） -->
      <div v-if="!isMultiProduct && results.dcf?.result?.details" class="card">
        <div class="card-title">📊 企业价值构成分析</div>
        <div ref="valueCompositionChart" class="chart"></div>
        <div class="value-composition-details">
          <div class="composition-item">
            <span class="composition-label">预测期现值（5年）</span>
            <span class="composition-value">{{ formatMoney(results.dcf.result.details.pv_forecasts) }}</span>
            <span class="composition-percent">{{ ((results.dcf.result.details.pv_forecasts / (results.dcf.result.details.pv_forecasts + results.dcf.result.details.pv_terminal)) * 100).toFixed(1) }}%</span>
          </div>
          <div class="composition-item">
            <span class="composition-label">终值现值（永续增长）</span>
            <span class="composition-value">{{ formatMoney(results.dcf.result.details.pv_terminal) }}</span>
            <span class="composition-percent">{{ ((results.dcf.result.details.pv_terminal / (results.dcf.result.details.pv_forecasts + results.dcf.result.details.pv_terminal)) * 100).toFixed(1) }}%</span>
          </div>
        </div>
      </div>

      <!-- 综合估值对比（仅单产品模式） -->
      <div v-if="!isMultiProduct && hasMultipleValuations" class="card">
        <div class="card-title">📊 综合估值对比</div>
        <div ref="comparisonChart" class="chart"></div>
        <div class="recommendation">
          <div class="rec-label">推荐估值（中位数）</div>
          <div class="rec-value">{{ formatMoney(getRecommendedValue()) }}</div>
          <div class="rec-range">估值区间: {{ getValueRange() }}</div>
        </div>
      </div>

      <!-- 情景分析（仅单产品模式） -->
      <div v-if="!isMultiProduct" class="card">
        <div class="card-title">📈 情景分析</div>
        <div ref="scenarioChart" class="chart"></div>
      </div>

      <!-- 敏感性分析（仅单产品模式） -->
      <div v-if="!isMultiProduct" class="card">
        <div class="card-title">📊 参数敏感性分析</div>
        <div ref="tornadoChart" class="chart"></div>

        <!-- 敏感性参数表格 -->
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

      <!-- 压力测试（仅单产品模式） -->
      <div v-if="!isMultiProduct" class="card">
        <div class="card-title">⚠️ 压力测试结果</div>
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
              <tr v-for="(test, idx) in sortedStressTests" :key="idx"
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

      <!-- 蒙特卡洛模拟（仅单产品模式） -->
      <div v-if="!isMultiProduct && monteCarloData" class="card">
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

      <!-- 历史记录 -->
      <div class="card">
        <div class="card-title">
          📋 历史记录
          <button @click="loadHistory" class="btn-refresh" :disabled="loadingHistory" type="button">
            {{ loadingHistory ? '加载中...' : '🔄 刷新' }}
          </button>
        </div>
        <div v-if="history.length === 0 && !loadingHistory" class="no-history">
          <p>暂无历史记录</p>
        </div>
        <div v-else-if="history.length > 0" class="history-list">
          <div v-for="item in history" :key="item.id" class="history-item" @click="loadHistoryItem(item.id)">
            <div class="history-header">
              <span class="history-company">{{ item.company_name }}</span>
              <span class="history-date">{{ formatDate(item.created_at) }}</span>
              <span class="history-industry">{{ item.industry }}</span>
            </div>
            <div class="history-values">
              <span v-if="item.dcf_value">DCF: {{ formatMoney(item.dcf_value * 10000) }}</span>
              <span v-if="item.pe_value">P/E: {{ formatMoney(item.pe_value * 10000) }}</span>
              <span v-if="item.ev_value">EV: {{ formatMoney(item.ev_value * 10000) }}</span>
            </div>
          </div>
        </div>
      </div>
      </template>
    </template>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, nextTick, computed } from 'vue'
import * as echarts from 'echarts'
import axios from 'axios'

const results = ref<any>(null)
const company = ref<any>(null)
const relativeChart = ref<HTMLElement>()
const comparisonChart = ref<HTMLElement>()
const valueCompositionChart = ref<HTMLElement>()
const scenarioChart = ref<HTMLElement>()
const tornadoChart = ref<HTMLElement>()
const monteCarloChart = ref<HTMLElement>()
const productContributionChart = ref<HTMLElement>()
const consolidatedCashFlowChart = ref<HTMLElement>()

const monteCarloData = computed(() => {
  return results.value?.stress?.report?.monte_carlo || null
})

// 检查是否为多产品估值模式
const isMultiProduct = computed(() => {
  return results.value?.valuationMode === 'multi'
})

// 多产品估值数据
const multiProductData = computed(() => {
  if (!isMultiProduct.value) return null
  return results.value?.multiProduct || null
})

// 修正后的企业价值（累加所有产品的enterprise_value，而不是使用后端返回的total_enterprise_value）
const correctedTotalEnterpriseValue = computed(() => {
  if (!multiProductData.value?.product_results) return 0
  return multiProductData.value.product_results.reduce((sum: number, p: any) => sum + (p.enterprise_value || 0), 0)
})

// 修正后的股权价值
const correctedTotalEquityValue = computed(() => {
  if (!multiProductData.value) return 0
  const company = results.value?.company || {}
  const netDebt = (company.total_debt || 0) - (company.cash_and_equivalents || 0)
  return correctedTotalEnterpriseValue.value - netDebt
})

// 历史记录
const history = ref<any[]>([])
const loadingHistory = ref(false)

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

const sortedStressTests = computed(() => {
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
      return match ? parseFloat(match[1] || '0') : 0
    }
    const pctA = getShockPct(a.scenario_description)
    const pctB = getShockPct(b.scenario_description)
    return pctB - pctA  // 降序排列
  })
})

const hasMultipleValuations = computed(() => {
  const hasRelative = results.value?.relative && Object.keys(results.value.relative.results || {}).length > 0
  const hasDCF = results.value?.dcf?.result?.value
  return hasRelative && hasDCF
})

onMounted(async () => {
  console.log('=== ValuationResult onMounted 开始 ===')

  // 先检查sessionStorage中所有的keys
  console.log('sessionStorage所有键值:', Object.keys(sessionStorage))

  const data = sessionStorage.getItem('valuationResults')
  console.log('ValuationResult onMounted - sessionStorage数据:', data)
  console.log('ValuationResult onMounted - 数据长度:', data?.length)
  console.log('ValuationResult onMounted - 数据前200字符:', data?.substring(0, 200))

  if (data) {
    try {
      const parsed = JSON.parse(data)
      console.log('ValuationResult onMounted - 解析后的数据:', parsed)
      console.log('ValuationResult onMounted - DCF数据:', parsed.dcf)
      console.log('ValuationResult onMounted - 相对估值数据:', parsed.relative)
      console.log('ValuationResult onMounted - 所有数据键:', Object.keys(parsed))

      // 详细检查相对估值数据
      console.log('ValuationResult onMounted - 相对估值详细数据:', parsed.relative)
      console.log('ValuationResult onMounted - parsed.relative类型:', typeof parsed.relative)
      console.log('ValuationResult onMounted - parsed.rel是否存在:', 'relative' in parsed)

      // 检查相对估值错误信息
      if (parsed.relativeError) {
        console.error('❌ 相对估值失败信息:', parsed.relativeError)
        console.error('错误消息:', parsed.relativeError.message)
        if (parsed.relativeError.response) {
          console.error('API响应:', parsed.relativeError.response)
        }
      }

      // 检查是否有comparables但没有relative
      if (parsed.hasComparables && !parsed.relative) {
        console.warn('⚠️ 有可比公司数据但无相对估值结果')
        console.warn('comparables:', parsed.comparables)
        if (parsed.noComparablesReason) {
          console.warn('跳过原因:', parsed.noComparablesReason)
        }
      }

      if (parsed.relative) {
        console.log('ValuationResult onMounted - parsed.relative结构:', JSON.stringify(parsed.relative, null, 2))
        if (parsed.relative.result) {
          console.log('ValuationResult onMounted - parsed.relative.result:', parsed.relative.result)
          console.log('ValuationResult onMounted - parsed.relative.result.results:', parsed.relative.result.results)
        }
      }

      results.value = parsed
      company.value = parsed.company

      await nextTick()
      if (isMultiProduct.value) {
        initMultiProductCharts()
        // 多产品模式下也要初始化相对估值图表
        if (results.value?.relative?.results) {
          console.log('✅ 多产品模式：初始化相对估值图表')
          initRelativeChart()
        } else {
          console.log('ℹ️ 多产品模式：无相对估值数据，跳过图表初始化')
        }
      } else {
        initCharts()
      }
    } catch (parseErr) {
      console.error('ValuationResult onMounted - JSON解析失败:', parseErr)
    }
  } else {
    console.error('ValuationResult onMounted - sessionStorage中没有valuationResults数据')
    console.error('ValuationResult onMounted - 请检查是否通过估值页面跳转而来')
  }

  console.log('=== ValuationResult onMounted 结束 ===')

  // 加载历史记录
  loadHistory()
})

// 初始化相对估值图表（独立函数，供多产品和单产品模式共用）
const initRelativeChart = () => {
  if (!results.value?.relative?.results || !relativeChart.value) return

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

const initCharts = () => {
  if (!results.value) return

  // 初始化相对估值图表
  initRelativeChart()

  // 初始化价值构成图表（单产品模式）
  if (valueCompositionChart.value && results.value.dcf?.result?.details) {
    const chart = echarts.init(valueCompositionChart.value)
    const details = results.value.dcf.result.details
    const pvForecasts = details.pv_forecasts || 0
    const pvTerminal = details.pv_terminal || 0

    chart.setOption({
      title: { text: '企业价值构成', left: 'center' },
      tooltip: {
        trigger: 'item',
        formatter: '{a} <br/>{b}: {c} 万元 ({d}%)'
      },
      legend: {
        orient: 'vertical',
        left: 'left'
      },
      series: [
        {
          name: '企业价值',
          type: 'pie',
          radius: ['40%', '70%'],
          avoidLabelOverlap: false,
          label: {
            show: true,
            formatter: '{b}: {d}%'
          },
          emphasis: {
            label: {
              show: true,
              fontSize: '16',
              fontWeight: 'bold'
            }
          },
          data: [
            {
              value: pvForecasts,
              name: '预测期现值（5年）',
              itemStyle: { color: '#667eea' }
            },
            {
              value: pvTerminal,
              name: '终值现值（永续增长）',
              itemStyle: { color: '#764ba2' }
            }
          ]
        }
      ]
    })
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
          params.push(getParameterName(paramName))
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
      xAxis: { type: 'category', data: bins.map((b: number) => b.toFixed(1)), name: '估值（亿元）' },
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

// 初始化多产品估值图表
const initMultiProductCharts = () => {
  if (!multiProductData.value) return

  // 初始化产品价值贡献饼图
  if (productContributionChart.value) {
    const chart = echarts.init(productContributionChart.value)
    const contributions = multiProductData.value.product_contribution || []

    const data = contributions.map((c: any) => ({
      name: c.product,
      value: c.contribution_pct
    }))

    chart.setOption({
      title: { text: '产品价值贡献占比', left: 'center' },
      tooltip: {
        trigger: 'item',
        formatter: '{b}: {c}%'
      },
      legend: {
        orient: 'vertical',
        left: 'left',
        top: 'middle'
      },
      series: [{
        type: 'pie',
        radius: ['40%', '70%'],
        center: ['60%', '50%'],
        data: data,
        emphasis: {
          itemStyle: {
            shadowBlur: 10,
            shadowOffsetX: 0,
            shadowColor: 'rgba(0, 0, 0, 0.5)'
          }
        },
        label: {
          formatter: '{b}\n{c}%'
        }
      }]
    })
  }

  // 初始化合并现金流图表
  if (consolidatedCashFlowChart.value) {
    const chart = echarts.init(consolidatedCashFlowChart.value)
    const forecasts = multiProductData.value.consolidated_fcf_forecasts || []

    const years = forecasts.map((f: any) => `第${f.year}年`)
    const revenues = forecasts.map((f: any) => (f.revenue / 10000).toFixed(2))
    const fcfs = forecasts.map((f: any) => (f.fcf / 10000).toFixed(2))

    chart.setOption({
      title: { text: '合并现金流预测', left: 'center' },
      tooltip: {
        trigger: 'axis',
        formatter: (params: any) => {
          let result = params[0].name + '<br/>'
          params.forEach((param: any) => {
            result += `${param.seriesName}: ${param.value} 亿元<br/>`
          })
          return result
        }
      },
      legend: {
        data: ['收入', '自由现金流'],
        top: 30
      },
      xAxis: {
        type: 'category',
        data: years
      },
      yAxis: {
        type: 'value',
        name: '金额（亿元）'
      },
      series: [
        {
          name: '收入',
          type: 'bar',
          data: revenues,
          itemStyle: { color: '#5470c6' }
        },
        {
          name: '自由现金流',
          type: 'line',
          data: fcfs,
          itemStyle: { color: '#91cc75' },
          lineStyle: { width: 3 }
        }
      ]
    })
  }

  // 初始化综合估值对比图表（多产品模式）
  if (comparisonChart.value && results.value?.relative?.results) {
    const chart = echarts.init(comparisonChart.value)
    const methods: string[] = []
    const values: number[] = []

    // 添加相对估值方法
    for (const [method, result] of Object.entries(results.value.relative.results)) {
      methods.push(getMethodName(method))
      values.push(((result as any).value || 0) / 10000)
    }

    // 添加多产品DCF
    if (correctedTotalEnterpriseValue.value) {
      methods.push('多产品DCF')
      values.push(correctedTotalEnterpriseValue.value / 10000)
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

  // 响应式调整
  window.addEventListener('resize', () => {
    if (productContributionChart.value) {
      const c = echarts.getInstanceByDom(productContributionChart.value)
      c?.resize()
    }
    if (consolidatedCashFlowChart.value) {
      const c = echarts.getInstanceByDom(consolidatedCashFlowChart.value)
      c?.resize()
    }
    if (relativeChart.value) {
      const c = echarts.getInstanceByDom(relativeChart.value)
      c?.resize()
    }
    if (comparisonChart.value) {
      const c = echarts.getInstanceByDom(comparisonChart.value)
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

// 多产品模式下的推荐估值
const getRecommendedValueForMultiProduct = () => {
  const values: number[] = []

  if (results.value?.relative?.results) {
    for (const result of Object.values(results.value.relative.results)) {
      values.push((result as any).value || 0)
    }
  }

  if (correctedTotalEnterpriseValue.value) {
    values.push(correctedTotalEnterpriseValue.value)
  }

  if (values.length === 0) return 0

  // 返回中位数
  values.sort((a, b) => a - b)
  return values[Math.floor(values.length / 2)]
}

// 多产品模式下的估值区间
const getValueRangeForMultiProduct = () => {
  const values: number[] = []

  if (results.value?.relative?.results) {
    for (const result of Object.values(results.value.relative.results)) {
      values.push((result as any).value || 0)
    }
  }

  if (correctedTotalEnterpriseValue.value) {
    values.push(correctedTotalEnterpriseValue.value)
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

// 加载历史记录
const loadHistory = async () => {
  loadingHistory.value = true
  try {
    const response = await axios.get('http://localhost:8000/api/history?limit=10')
    if (response.data.success) {
      history.value = response.data.history
    }
  } catch (err: any) {
    console.error('加载历史记录失败:', err)
  } finally {
    loadingHistory.value = false
  }
}

// 加载历史记录项
const loadHistoryItem = async (id: number) => {
  try {
    const response = await axios.get(`http://localhost:8000/api/history/${id}`)
    if (response.data.success) {
      const item = response.data.history
      // 存储到sessionStorage并跳转
      sessionStorage.setItem('valuationResults', JSON.stringify(item))
      results.value = item
      company.value = { name: item.company_name, industry: item.industry }
      await nextTick()
      initCharts()
    }
  } catch (err: any) {
    console.error('加载历史记录项失败:', err)
  }
}

// 格式化日期
const formatDate = (dateStr: string) => {
  if (!dateStr) return '--'
  const date = new Date(dateStr)
  return date.toLocaleDateString('zh-CN', { year: 'numeric', month: '2-digit', day: '2-digit' })
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

.no-data .hint {
  color: #666;
  font-size: 0.9em;
  margin: 15px 0;
}

.no-data .error-list {
  text-align: left;
  max-width: 400px;
  margin: 0 auto;
  color: #555;
  font-size: 0.85em;
}

.no-data .error-list li {
  margin: 8px 0;
  padding-left: 20px;
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

.sensitivity-table-container {
  overflow-x: auto;
  margin-top: 20px;
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

/* 历史记录样式 */
.btn-refresh {
  background: #667eea;
  color: white;
  border: none;
  padding: 6px 12px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 0.9em;
  transition: all 0.3s;
}

.btn-refresh:hover:not(:disabled) {
  background: #5568d3;
}

.btn-refresh:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.no-history {
  text-align: center;
  padding: 40px 20px;
  background: #f8f9fa;
  border-radius: 8px;
  color: #666;
}

.history-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.history-item {
  background: #f8f9fa;
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  padding: 15px;
  cursor: pointer;
  transition: all 0.2s;
}

.history-item:hover {
  background: #e8f4ff;
  border-color: #667eea;
}

.history-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 10px;
}

.history-company {
  font-weight: 600;
  color: #333;
  font-size: 1.05em;
}

.history-date {
  font-size: 0.85em;
  color: #666;
}

.history-industry {
  font-size: 0.9em;
  color: #888;
  background: #f0f0f0;
  padding: 3px 8px;
  border-radius: 4px;
}

.history-values {
  display: flex;
  gap: 15px;
  flex-wrap: wrap;
}

.history-values span {
  font-size: 0.9em;
  color: #555;
  background: white;
  padding: 4px 10px;
  border-radius: 4px;
  border: 1px solid #e0e0e0;
}

/* 多产品估值样式 */
.multi-product-section {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.product-contribution-list {
  margin-top: 20px;
  display: flex;
  flex-direction: column;
  gap: 15px;
}

.contribution-item {
  background: #f8f9fa;
  padding: 15px;
  border-radius: 8px;
}

.contribution-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 8px;
}

.contribution-name {
  font-weight: 600;
  color: #333;
  font-size: 1.05em;
}

.contribution-percent {
  font-size: 1.2em;
  color: #667eea;
  font-weight: bold;
}

.contribution-bar {
  height: 12px;
  background: #e0e0e0;
  border-radius: 6px;
  overflow: hidden;
}

.contribution-fill {
  height: 100%;
  background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
  border-radius: 6px;
  transition: width 0.5s ease;
}

.product-table-container {
  overflow-x: auto;
  margin-top: 15px;
}

.product-table {
  width: 100%;
  border-collapse: collapse;
  background: white;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.product-table thead {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
}

.product-table th {
  padding: 12px 16px;
  text-align: left;
  font-weight: 600;
  font-size: 0.95em;
}

.product-table tbody tr {
  border-bottom: 1px solid #f0f0f0;
  transition: background 0.2s;
}

.product-table tbody tr:hover {
  background: #f8f9fa;
}

.product-table tbody tr:last-child {
  border-bottom: none;
}

.product-table td {
  padding: 12px 16px;
  color: #333;
}

.product-name {
  font-weight: 600;
  color: #555;
}

.value-contribution {
  font-weight: 700;
  color: #667eea;
}

.enterprise-value {
  font-weight: 700;
  color: #667eea;
}

.info-note {
  margin-top: 15px;
  padding: 12px 16px;
  background: #e8f4ff;
  border-left: 4px solid #667eea;
  border-radius: 6px;
  font-size: 0.9em;
  color: #555;
}

.value-composition-details {
  margin-top: 20px;
  padding: 16px;
  background: #f8f9fa;
  border-radius: 8px;
}

.composition-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 12px 0;
  border-bottom: 1px solid #e0e0e0;
}

.composition-item:last-child {
  border-bottom: none;
}

.composition-label {
  flex: 1;
  font-weight: 500;
  color: #333;
}

.composition-value {
  flex: 1;
  text-align: right;
  font-weight: 600;
  color: #667eea;
}

.composition-percent {
  flex: 0 0 80px;
  text-align: right;
  font-weight: 700;
  font-size: 1.1em;
  color: #764ba2;
}

.product-cashflow-section {
  margin-bottom: 25px;
  padding: 20px;
  background: #f8f9fa;
  border-radius: 8px;
  border-left: 4px solid #667eea;
}

.product-section-title {
  margin: 0 0 15px 0;
  color: #333;
  font-size: 1.1em;
  padding-bottom: 10px;
  border-bottom: 2px solid #e0e0e0;
}

.product-cashflow-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 12px;
}

.forecast-item {
  background: white;
  padding: 12px;
  border-radius: 6px;
  border: 1px solid #e0e0e0;
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.forecast-year {
  font-weight: 600;
  color: #667eea;
  font-size: 0.95em;
}

.forecast-revenue,
.forecast-fcf,
.forecast-growth {
  font-size: 0.85em;
  color: #555;
}

.forecast-growth {
  color: #91cc75;
  font-weight: 500;
}
</style>
