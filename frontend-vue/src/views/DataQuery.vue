<template>
  <div class="data-query">
    <div class="header">
      <h1>📊 数据查询</h1>
      <p>查询和管理估值历史记录，支持数据对比和趋势分析</p>
    </div>

    <!-- 筛选条件区域 -->
    <div class="card filters-card">
      <div class="filter-title">🔍 筛选条件</div>
      <div class="filters-grid">
        <div class="filter-item">
          <label>公司名称</label>
          <input
            v-model="filters.companyName"
            type="text"
            placeholder="输入公司名称（支持模糊匹配）"
            @input="filterRecords"
          />
        </div>
        <div class="filter-item">
          <label>开始日期</label>
          <input
            v-model="filters.startDate"
            type="date"
            @change="filterRecords"
          />
        </div>
        <div class="filter-item">
          <label>结束日期</label>
          <input
            v-model="filters.endDate"
            type="date"
            @change="filterRecords"
          />
        </div>
        <div class="filter-item">
          <label>行业</label>
          <select
            v-model="filters.industry"
            @change="filterRecords"
          >
            <option value="">全部行业</option>
            <option value="制造业">制造业</option>
            <option value="金融业">金融业</option>
            <option value="科技业">科技业</option>
            <option value="医疗业">医疗业</option>
            <option value="消费品">消费品</option>
            <option value="能源业">能源业</option>
            <option value="其他">其他</option>
          </select>
        </div>
        <div class="filter-item">
          <label>估值方法</label>
          <select
            v-model="filters.method"
            @change="filterRecords"
          >
            <option value="">全部方法</option>
            <option value="DCF">DCF绝对估值</option>
            <option value="PE">市盈率法</option>
            <option value="PB">市净率法</option>
            <option value="PS">市销率法</option>
          </select>
        </div>
        <div class="filter-actions">
          <button @click="resetFilters" class="btn-secondary">重置筛选</button>
          <button @click="exportData" class="btn-export">📥 导出数据</button>
        </div>
      </div>
    </div>

    <!-- 历史记录统计 -->
    <div class="stats-summary">
      <div class="stat-item">
        <div class="stat-label">总记录数</div>
        <div class="stat-value">{{ filteredRecords.length }}</div>
      </div>
      <div class="stat-item">
        <div class="stat-label">筛选结果</div>
        <div class="stat-value">{{ filteredRecords.length }} 条</div>
      </div>
      <div class="stat-item">
        <div class="stat-label">平均估值</div>
        <div class="stat-value">{{ formatMoney(averageValuation) }}</div>
      </div>
      <div class="stat-item">
        <div class="stat-label">最早记录</div>
        <div class="stat-value">{{ earliestDate }}</div>
      </div>
    </div>

    <!-- 历史记录列表 -->
    <div class="card records-card">
      <div class="records-header">
        <div class="records-title">📋 估值历史记录</div>
        <div class="records-count">共 {{ filteredRecords.length }} 条记录</div>
      </div>

      <div v-if="filteredRecords.length === 0" class="no-data">
        <p>暂无历史记录</p>
        <p class="hint">完成估值后会自动保存到历史记录中</p>
        <button @click="$router.push('/valuation')" class="btn-primary">
          开始新的估值
        </button>
      </div>

      <div v-else class="records-grid">
        <div
          v-for="record in paginatedRecords"
          :key="record.id"
          class="record-card"
        >
          <div class="record-header">
            <div class="record-company">{{ record.companyName }}</div>
            <div class="record-date">{{ formatDate(record.valuationDate) }}</div>
          </div>
          <div class="record-details">
            <div class="record-item">
              <span class="label">行业:</span>
              <span class="value">{{ record.industry }}</span>
            </div>
            <div class="record-item">
              <span class="label">阶段:</span>
              <span class="value">{{ record.stage }}</span>
            </div>
            <div class="record-item">
              <span class="label">方法:</span>
              <span class="value">{{ record.valuationMethod }}</span>
            </div>
            <div class="record-valuation">
              <div class="valuation-label">估值结果</div>
              <div class="valuation-value">{{ formatMoney(record.valuation) }}</div>
            </div>
          </div>
          <div class="record-actions">
            <button @click="viewDetails(record)" class="btn-view">
              📄 查看详情
            </button>
            <button @click="confirmDelete(record)" class="btn-delete">
              🗑️ 删除
            </button>
          </div>
        </div>
      </div>

      <!-- 分页控制 -->
      <div v-if="filteredRecords.length > pageSize" class="pagination">
        <button
          @click="currentPage--"
          :disabled="currentPage === 1"
          class="btn-page"
        >
          上一页
        </button>
        <span class="page-info">第 {{ currentPage }} 页，共 {{ totalPages }} 页</span>
        <button
          @click="currentPage++"
          :disabled="currentPage === totalPages"
          class="btn-page"
        >
          下一页
        </button>
      </div>
    </div>

    <!-- 详情弹窗 -->
    <div v-if="showDetailModal" class="modal-overlay" @click="closeDetailModal">
      <div class="modal-content" @click.stop>
        <div class="modal-header">
          <h2>📄 估值详情</h2>
          <button @click="closeDetailModal" class="btn-close">✕</button>
        </div>
        <div class="modal-body" v-if="selectedRecord">
          <div class="detail-section">
            <h3>基本信息</h3>
            <div class="detail-grid">
              <div class="detail-item">
                <span class="label">公司名称:</span>
                <span class="value">{{ selectedRecord.companyName }}</span>
              </div>
              <div class="detail-item">
                <span class="label">估值日期:</span>
                <span class="value">{{ formatDate(selectedRecord.valuationDate) }}</span>
              </div>
              <div class="detail-item">
                <span class="label">行业:</span>
                <span class="value">{{ selectedRecord.industry }}</span>
              </div>
              <div class="detail-item">
                <span class="label">发展阶段:</span>
                <span class="value">{{ selectedRecord.stage }}</span>
              </div>
              <div class="detail-item">
                <span class="label">估值方法:</span>
                <span class="value">{{ selectedRecord.valuationMethod }}</span>
              </div>
            </div>
          </div>

          <div class="detail-section">
            <h3>财务数据</h3>
            <div class="detail-grid">
              <div class="detail-item">
                <span class="label">营业收入:</span>
                <span class="value">{{ formatMoney(selectedRecord.revenue) }}</span>
              </div>
              <div class="detail-item">
                <span class="label">净利润:</span>
                <span class="value">{{ formatMoney(selectedRecord.netIncome) }}</span>
              </div>
              <div class="detail-item">
                <span class="label">总资产:</span>
                <span class="value">{{ formatMoney(selectedRecord.totalAssets) }}</span>
              </div>
              <div class="detail-item">
                <span class="label">净资产:</span>
                <span class="value">{{ formatMoney(selectedRecord.netAssets) }}</span>
              </div>
            </div>
          </div>

          <div class="detail-section">
            <h3>估值参数</h3>
            <div class="detail-grid">
              <div class="detail-item">
                <span class="label">增长率:</span>
                <span class="value">{{ (selectedRecord.growthRate * 100).toFixed(1) }}%</span>
              </div>
              <div class="detail-item">
                <span class="label">WACC:</span>
                <span class="value">{{ (selectedRecord.wacc * 100).toFixed(1) }}%</span>
              </div>
              <div class="detail-item">
                <span class="label">营业利润率:</span>
                <span class="value">{{ (selectedRecord.operatingMargin * 100).toFixed(1) }}%</span>
              </div>
            </div>
          </div>

          <div class="detail-section">
            <h3>估值结果</h3>
            <div class="valuation-result">
              <div class="result-main">
                <div class="result-label">估值结果</div>
                <div class="result-value">{{ formatMoney(selectedRecord.valuation) }}</div>
              </div>
              <div class="result-details">
                <div v-if="selectedRecord.dcfValuation" class="result-item">
                  <span class="method">DCF估值:</span>
                  <span class="value">{{ formatMoney(selectedRecord.dcfValuation) }}</span>
                </div>
                <div v-if="selectedRecord.peValuation" class="result-item">
                  <span class="method">PE估值:</span>
                  <span class="value">{{ formatMoney(selectedRecord.peValuation) }}</span>
                </div>
                <div v-if="selectedRecord.pbValuation" class="result-item">
                  <span class="method">PB估值:</span>
                  <span class="value">{{ formatMoney(selectedRecord.pbValuation) }}</span>
                </div>
              </div>
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button @click="closeDetailModal" class="btn-primary">关闭</button>
        </div>
      </div>
    </div>

    <!-- 删除确认弹窗 -->
    <div v-if="showDeleteModal" class="modal-overlay" @click="closeDeleteModal">
      <div class="modal-content modal-small" @click.stop>
        <div class="modal-header">
          <h2>⚠️ 确认删除</h2>
        </div>
        <div class="modal-body">
          <p>确定要删除这条估值记录吗？</p>
          <p class="delete-warning">删除后将无法恢复</p>
          <div v-if="recordToDelete" class="record-info">
            <strong>{{ recordToDelete.companyName }}</strong><br>
            <small>{{ formatDate(recordToDelete.valuationDate) }}</small><br>
            <small>估值: {{ formatMoney(recordToDelete.valuation) }}</small>
          </div>
        </div>
        <div class="modal-footer">
          <button @click="closeDeleteModal" class="btn-secondary">取消</button>
          <button @click="deleteRecord" class="btn-delete">确认删除</button>
        </div>
      </div>
    </div>

    <!-- 趋势图表 -->
    <div class="card trend-card">
      <div class="trend-header">
        <div class="trend-title">📈 估值趋势分析</div>
        <div class="trend-controls">
          <select v-model="trendPeriod" @change="updateTrendChart">
            <option value="30">最近30条记录</option>
            <option value="90">最近90条记录</option>
            <option value="180">最近180条记录</option>
            <option value="all">全部记录</option>
          </select>
        </div>
      </div>
      <div ref="trendChart" class="chart"></div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, watch } from 'vue'
import * as echarts from 'echarts'

// 数据类型定义
interface ValuationRecord {
  id: number
  companyName: string
  valuationDate: string
  industry: string
  stage: string
  valuationMethod: string
  valuation: number
  revenue: number
  netIncome: number
  totalAssets: number
  netAssets: number
  growthRate: number
  wacc: number
  operatingMargin: number
  dcfValuation?: number
  peValuation?: number
  pbValuation?: number
  psValuation?: number
}

// 状态管理
const records = ref<ValuationRecord[]>([])
const filters = ref({
  companyName: '',
  startDate: '',
  endDate: '',
  industry: '',
  method: ''
})

const currentPage = ref(1)
const pageSize = 12
const showDetailModal = ref(false)
const showDeleteModal = ref(false)
const selectedRecord = ref<ValuationRecord | null>(null)
const recordToDelete = ref<ValuationRecord | null>(null)
const trendPeriod = ref('30')
const trendChart = ref<HTMLElement>()

// 计算属性
const filteredRecords = computed(() => {
  let result = [...records.value]

  // 按时间倒序排列
  result.sort((a, b) => new Date(b.valuationDate).getTime() - new Date(a.valuationDate).getTime())

  // 应用筛选条件
  if (filters.value.companyName) {
    const keyword = filters.value.companyName.toLowerCase()
    result = result.filter(record =>
      record.companyName.toLowerCase().includes(keyword)
    )
  }

  if (filters.value.startDate) {
    result = result.filter(record =>
      new Date(record.valuationDate) >= new Date(filters.value.startDate)
    )
  }

  if (filters.value.endDate) {
    result = result.filter(record =>
      new Date(record.valuationDate) <= new Date(filters.value.endDate)
    )
  }

  if (filters.value.industry) {
    result = result.filter(record => record.industry === filters.value.industry)
  }

  if (filters.value.method) {
    result = result.filter(record => record.valuationMethod === filters.value.method)
  }

  return result
})

const paginatedRecords = computed(() => {
  const start = (currentPage.value - 1) * pageSize
  const end = start + pageSize
  return filteredRecords.value.slice(start, end)
})

const totalPages = computed(() => {
  return Math.ceil(filteredRecords.value.length / pageSize)
})

const averageValuation = computed(() => {
  if (filteredRecords.value.length === 0) return 0
  const sum = filteredRecords.value.reduce((acc, record) => acc + record.valuation, 0)
  return sum / filteredRecords.value.length
})

const earliestDate = computed(() => {
  if (filteredRecords.value.length === 0) return '--'
  const dates = filteredRecords.value.map(record => new Date(record.valuationDate))
  const minDate = new Date(Math.min.apply(null, dates))
  return minDate.toLocaleDateString('zh-CN')
})

// 方法
const formatMoney = (value: number) => {
  if (!value) return '--'
  return (value / 10000).toFixed(2) + ' 亿元'
}

const formatDate = (dateStr: string) => {
  const date = new Date(dateStr)
  return date.toLocaleDateString('zh-CN')
}

const filterRecords = () => {
  currentPage.value = 1
}

const resetFilters = () => {
  filters.value = {
    companyName: '',
    startDate: '',
    endDate: '',
    industry: '',
    method: ''
  }
  currentPage.value = 1
}

const viewDetails = (record: ValuationRecord) => {
  selectedRecord.value = record
  showDetailModal.value = true
}

const closeDetailModal = () => {
  showDetailModal.value = false
  selectedRecord.value = null
}

const confirmDelete = (record: ValuationRecord) => {
  recordToDelete.value = record
  showDeleteModal.value = true
}

const closeDeleteModal = () => {
  showDeleteModal.value = false
  recordToDelete.value = null
}

const deleteRecord = async () => {
  if (recordToDelete.value) {
    try {
      // 调用API删除记录
      const response = await fetch(`http://localhost:8000/api/history/${recordToDelete.value.id}`, {
        method: 'DELETE'
      })

      if (response.ok) {
        // 从本地状态中移除记录
        records.value = records.value.filter(
          record => record.id !== recordToDelete.value.id
        )

        // 显示成功消息
        alert(`已删除 "${recordToDelete.value.companyName}" 的估值记录`)

        closeDeleteModal()
        updateTrendChart()
      } else {
        throw new Error('删除失败')
      }
    } catch (error) {
      console.error('删除记录失败:', error)
      // 如果API调用失败，只从本地状态中移除
      records.value = records.value.filter(
        record => record.id !== recordToDelete.value.id
      )
      alert(`已从本地列表中删除 "${recordToDelete.value.companyName}" 的估值记录`)
      closeDeleteModal()
      updateTrendChart()
    }
  }
}

const exportData = () => {
  if (filteredRecords.value.length === 0) {
    alert('没有数据可导出')
    return
  }

  // 创建CSV数据
  const headers = ['公司名称', '估值日期', '行业', '阶段', '估值方法', '估值结果(亿元)', '营业收入(亿元)', '净利润(亿元)']
  const rows = filteredRecords.value.map(record => [
    record.companyName,
    formatDate(record.valuationDate),
    record.industry,
    record.stage,
    record.valuationMethod,
    (record.valuation / 10000).toFixed(2),
    (record.revenue / 10000).toFixed(2),
    (record.netIncome / 10000).toFixed(2)
  ])

  const csvContent = [headers, ...rows].map(row => row.join(',')).join('\\n')
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)

  const link = document.createElement('a')
  link.setAttribute('href', url)
  link.setAttribute('download', `估值历史记录_${new Date().toISOString().slice(0, 10)}.csv`)
  link.click()

  URL.revokeObjectURL(url)
}

const updateTrendChart = () => {
  if (!trendChart.value) return

  const chart = echarts.init(trendChart.value)

  // 根据选择的周期获取数据
  let trendData = [...records.value]
  if (trendPeriod.value !== 'all') {
    trendData = records.value.slice(0, parseInt(trendPeriod.value))
  }

  // 按时间排序
  trendData.sort((a, b) => new Date(a.valuationDate).getTime() - new Date(b.valuationDate).getTime())

  const dates = trendData.map(record => formatDate(record.valuationDate))
  const valuations = trendData.map(record => record.valuation / 10000)

  chart.setOption({
    title: {
      text: '估值趋势分析',
      left: 'center'
    },
    tooltip: {
      trigger: 'axis',
      formatter: (params: any) => {
        const date = params[0].axisValue
        const value = params[0].value
        return `${date}<br/>估值: ${value.toFixed(2)} 亿元`
      }
    },
    xAxis: {
      type: 'category',
      data: dates,
      axisLabel: {
        rotate: 45
      }
    },
    yAxis: {
      type: 'value',
      name: '估值（亿元）'
    },
    series: [{
      type: 'line',
      data: valuations,
      smooth: true,
      areaStyle: {
        color: 'rgba(102, 126, 234, 0.3)'
      },
      itemStyle: {
        color: '#667eea'
      }
    }]
  })
}

// 加载历史记录
const loadHistoryRecords = async () => {
  try {
    const response = await fetch('http://localhost:8000/api/history?limit=100')
    const data = await response.json()

    if (data.status === 'success' && data.history) {
      // 将API数据转换为我们的格式
      records.value = data.history.map((item: any) => ({
        id: item.id || Date.now() + Math.random(),
        companyName: item.company_name || item.companyName || '未知公司',
        valuationDate: item.valuation_date || item.valuationDate || new Date().toISOString(),
        industry: item.industry || '其他',
        stage: item.stage || '成长期',
        valuationMethod: item.valuation_method || item.valuationMethod || 'DCF',
        valuation: item.valuation || item.total_valuation || 0,
        revenue: item.revenue || 0,
        netIncome: item.net_income || item.netIncome || 0,
        totalAssets: item.total_assets || item.totalAssets || 0,
        netAssets: item.net_assets || item.netAssets || 0,
        growthRate: item.growth_rate || item.growthRate || 0.1,
        wacc: item.wacc || 0.1,
        operatingMargin: item.operating_margin || item.operatingMargin || 0.15,
        dcfValuation: item.dcf_valuation || item.dcfValuation,
        peValuation: item.pe_valuation || item.peValuation,
        pbValuation: item.pb_valuation || item.pbValuation,
        psValuation: item.ps_valuation || item.psValuation
      }))
    } else {
      // 如果API没有返回数据，使用模拟数据
      console.warn('API未返回数据，使用模拟数据')
      records.value = generateMockRecords(15)
    }
  } catch (error) {
    console.error('加载历史记录失败:', error)
    // 如果API调用失败，使用模拟数据
    records.value = generateMockRecords(15)
  }
}

// 生成模拟历史记录数据
const generateMockRecords = (count: number): ValuationRecord[] => {
  const companies = ['科技公司A', '制造企业B', '金融服务C', '医疗健康D', '消费品牌E']
  const industries = ['科技业', '制造业', '金融业', '医疗业', '消费品']
  const stages = ['早期', '成长期', '成熟期', '上市公司']
  const methods = ['DCF', 'PE', 'PB', 'PS']

  const records: ValuationRecord[] = []
  const now = new Date()

  for (let i = 0; i < count; i++) {
    const date = new Date(now)
    date.setDate(date.getDate() - i * 7) // 每周一条记录

    const companyIndex = i % companies.length
    const baseValuation = 10000 + Math.random() * 50000 // 1-6亿元

    records.push({
      id: i + 1,
      companyName: companies[companyIndex],
      valuationDate: date.toISOString(),
      industry: industries[companyIndex],
      stage: stages[companyIndex],
      valuationMethod: methods[i % methods.length],
      valuation: baseValuation,
      revenue: baseValuation * 0.8,
      netIncome: baseValuation * 0.15,
      totalAssets: baseValuation * 2,
      netAssets: baseValuation * 1.2,
      growthRate: 0.1 + Math.random() * 0.2,
      wacc: 0.1 + Math.random() * 0.05,
      operatingMargin: 0.15 + Math.random() * 0.1,
      dcfValuation: baseValuation * (0.9 + Math.random() * 0.2),
      peValuation: baseValuation * (0.8 + Math.random() * 0.4),
      pbValuation: baseValuation * (0.7 + Math.random() * 0.6)
    })
  }

  return records
}

// 监听筛选条件变化
watch(trendPeriod, () => {
  updateTrendChart()
})

// 生命周期
onMounted(() => {
  loadHistoryRecords()
  // 延迟初始化图表，确保DOM已渲染
  setTimeout(() => {
    updateTrendChart()
  }, 100)
})
</script>

<style scoped>
.data-query {
  padding: 20px;
  max-width: 1400px;
  margin: 0 auto;
}

.header {
  text-align: center;
  margin-bottom: 30px;
  padding: 30px 20px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  border-radius: 12px;
  color: white;
}

.header h1 {
  font-size: 2.2em;
  margin-bottom: 10px;
}

.header p {
  font-size: 1.1em;
  opacity: 0.9;
}

/* 筛选卡片 */
.filters-card {
  margin-bottom: 20px;
}

.filter-title {
  font-size: 1.2em;
  font-weight: bold;
  margin-bottom: 20px;
  color: #667eea;
}

.filters-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 15px;
  margin-bottom: 15px;
}

.filter-item {
  display: flex;
  flex-direction: column;
}

.filter-item label {
  font-size: 0.9em;
  color: #666;
  margin-bottom: 5px;
  font-weight: 500;
}

.filter-item input,
.filter-item select {
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 6px;
  font-size: 0.9em;
}

.filter-actions {
  display: flex;
  gap: 10px;
  margin-top: 10px;
}

/* 统计摘要 */
.stats-summary {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 15px;
  margin-bottom: 20px;
}

.stat-item {
  background: white;
  padding: 15px;
  border-radius: 8px;
  border-left: 4px solid #667eea;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.stat-label {
  font-size: 0.85em;
  color: #666;
  margin-bottom: 5px;
}

.stat-value {
  font-size: 1.5em;
  font-weight: bold;
  color: #333;
}

/* 记录卡片 */
.records-card {
  margin-bottom: 20px;
}

.records-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.records-title {
  font-size: 1.3em;
  font-weight: bold;
  color: #333;
}

.records-count {
  font-size: 0.9em;
  color: #666;
}

.no-data {
  text-align: center;
  padding: 60px 20px;
  color: #999;
}

.no-data p {
  margin-bottom: 15px;
}

.hint {
  font-size: 0.9em;
  color: #aaa;
}

.records-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
  gap: 20px;
  margin-bottom: 20px;
}

.record-card {
  background: white;
  border-radius: 10px;
  padding: 20px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
  border: 1px solid #eee;
  transition: transform 0.2s, box-shadow 0.2s;
}

.record-card:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
}

.record-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 15px;
  padding-bottom: 10px;
  border-bottom: 1px solid #f0f0f0;
}

.record-company {
  font-size: 1.1em;
  font-weight: bold;
  color: #333;
}

.record-date {
  font-size: 0.9em;
  color: #999;
}

.record-details {
  margin-bottom: 15px;
}

.record-item {
  display: flex;
  margin-bottom: 8px;
  font-size: 0.9em;
}

.record-item .label {
  color: #666;
  min-width: 60px;
  margin-right: 10px;
}

.record-item .value {
  color: #333;
  font-weight: 500;
}

.record-valuation {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  padding: 12px;
  border-radius: 6px;
  margin-bottom: 15px;
}

.valuation-label {
  font-size: 0.85em;
  color: rgba(255, 255, 255, 0.8);
  margin-bottom: 5px;
}

.valuation-value {
  font-size: 1.3em;
  font-weight: bold;
  color: white;
}

.record-actions {
  display: flex;
  gap: 10px;
}

.btn-view,
.btn-delete {
  flex: 1;
  padding: 8px 12px;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  font-size: 0.9em;
  font-weight: 500;
  transition: all 0.2s;
}

.btn-view {
  background: #667eea;
  color: white;
}

.btn-view:hover {
  background: #5568d3;
}

.btn-delete {
  background: #f56c6c;
  color: white;
}

.btn-delete:hover {
  background: #e55a5a;
}

/* 分页 */
.pagination {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 15px;
  margin-top: 20px;
  padding-top: 20px;
  border-top: 1px solid #eee;
}

.btn-page {
  padding: 8px 16px;
  border: 1px solid #ddd;
  background: white;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.2s;
}

.btn-page:hover:not(:disabled) {
  background: #f0f0f0;
}

.btn-page:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.page-info {
  font-size: 0.9em;
  color: #666;
}

/* 趋势图表 */
.trend-card {
  margin-bottom: 20px;
}

.trend-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.trend-title {
  font-size: 1.2em;
  font-weight: bold;
  color: #333;
}

.trend-controls select {
  padding: 6px 12px;
  border: 1px solid #ddd;
  border-radius: 6px;
  font-size: 0.9em;
}

.chart {
  height: 400px;
  width: 100%;
}

/* 弹窗样式 */
.modal-overlay {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  justify-content: center;
  align-items: center;
  z-index: 1000;
}

.modal-content {
  background: white;
  border-radius: 12px;
  max-width: 800px;
  width: 90%;
  max-height: 80vh;
  overflow-y: auto;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
}

.modal-small {
  max-width: 400px;
}

.modal-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 20px;
  border-bottom: 1px solid #eee;
}

.modal-header h2 {
  margin: 0;
  color: #333;
}

.btn-close {
  background: none;
  border: none;
  font-size: 1.5em;
  cursor: pointer;
  color: #999;
}

.modal-body {
  padding: 20px;
}

.detail-section {
  margin-bottom: 25px;
}

.detail-section h3 {
  color: #667eea;
  margin-bottom: 15px;
  font-size: 1.1em;
}

.detail-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 15px;
}

.detail-item {
  display: flex;
  flex-direction: column;
}

.detail-item .label {
  font-size: 0.85em;
  color: #666;
  margin-bottom: 3px;
}

.detail-item .value {
  font-size: 1em;
  color: #333;
  font-weight: 500;
}

.valuation-result {
  background: #f8f9fa;
  padding: 15px;
  border-radius: 8px;
}

.result-main {
  margin-bottom: 15px;
  text-align: center;
}

.result-label {
  font-size: 0.9em;
  color: #666;
  margin-bottom: 5px;
}

.result-value {
  font-size: 2em;
  font-weight: bold;
  color: #667eea;
}

.result-details {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 10px;
}

.result-item {
  display: flex;
  justify-content: space-between;
  padding: 8px;
  background: white;
  border-radius: 4px;
  font-size: 0.9em;
}

.result-item .method {
  color: #666;
}

.result-item .value {
  color: #333;
  font-weight: 500;
}

.modal-footer {
  padding: 15px 20px;
  border-top: 1px solid #eee;
  display: flex;
  justify-content: flex-end;
  gap: 10px;
}

.delete-warning {
  color: #f56c6c;
  font-weight: 500;
  margin: 10px 0;
}

.record-info {
  background: #fff3cd;
  padding: 10px;
  border-radius: 6px;
  margin: 10px 0;
  font-size: 0.9em;
}

/* 通用按钮样式 */
.btn-primary,
.btn-secondary,
.btn-export {
  padding: 8px 16px;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  font-size: 0.9em;
  font-weight: 500;
  transition: all 0.2s;
}

.btn-primary {
  background: #667eea;
  color: white;
}

.btn-primary:hover {
  background: #5568d3;
}

.btn-secondary {
  background: #6c757d;
  color: white;
}

.btn-secondary:hover {
  background: #5a6268;
}

.btn-export {
  background: #28a745;
  color: white;
}

.btn-export:hover {
  background: #218838;
}

/* 响应式设计 */
@media (max-width: 768px) {
  .filters-grid {
    grid-template-columns: 1fr;
  }

  .stats-summary {
    grid-template-columns: repeat(2, 1fr);
  }

  .records-grid {
    grid-template-columns: 1fr;
  }

  .record-actions {
    flex-direction: column;
  }
}
</style>