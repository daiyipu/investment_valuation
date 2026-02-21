<template>
  <div class="history-page">
    <div class="header">
      <h1>📋 历史记录</h1>
      <p>查看历史估值分析记录</p>
    </div>

    <div class="card">
      <div class="card-actions">
        <button @click="loadHistory" class="btn-refresh" :disabled="loading">
          {{ loading ? '加载中...' : '🔄 刷新' }}
        </button>
      </div>

      <div v-if="history.length === 0 && !loading" class="no-history">
        <p>暂无历史记录</p>
        <p class="hint">进行估值分析后，记录将自动保存在这里</p>
        <router-link to="/valuation" class="btn-primary">开始估值</router-link>
      </div>

      <div v-else-if="history.length > 0" class="history-list">
        <div v-for="item in history" :key="item.id" class="history-item" @click="viewHistoryItem(item.id)">
          <div class="history-header">
            <span class="history-company">{{ item.company_name }}</span>
            <span class="history-date">{{ formatDate(item.created_at) }}</span>
          </div>
          <div class="history-meta">
            <span class="history-industry">{{ item.industry }}</span>
            <span class="history-stage">{{ item.stage }}</span>
          </div>
          <div class="history-values">
            <span v-if="item.dcf_value" class="value-item">
              DCF: {{ formatMoney(item.dcf_value) }}
            </span>
            <span v-if="item.pe_value" class="value-item">
              P/E: {{ formatMoney(item.pe_value) }}
            </span>
            <span v-if="item.ps_value" class="value-item">
              P/S: {{ formatMoney(item.ps_value) }}
            </span>
          </div>
        </div>
      </div>
    </div>

    <!-- 历史记录详情弹窗 -->
    <div v-if="selectedItem && showModal" class="modal-overlay" @click.self="closeModal">
      <div class="modal-content">
        <div class="modal-header">
          <h2>{{ selectedItem.company_name }} - 估值详情</h2>
          <button @click="closeModal" class="btn-close">×</button>
        </div>
        <div class="modal-body">
          <div class="detail-grid">
            <div class="detail-item">
              <span class="detail-label">行业</span>
              <span class="detail-value">{{ selectedItem.industry }}</span>
            </div>
            <div class="detail-item">
              <span class="detail-label">阶段</span>
              <span class="detail-value">{{ selectedItem.stage }}</span>
            </div>
            <div class="detail-item">
              <span class="detail-label">营业收入</span>
              <span class="detail-value">{{ formatMoney(selectedItem.revenue) }}</span>
            </div>
            <div class="detail-item">
              <span class="detail-label">创建时间</span>
              <span class="detail-value">{{ formatDateTime(selectedItem.created_at) }}</span>
            </div>
          </div>

          <div class="valuation-results">
            <h3>估值结果</h3>
            <div v-if="selectedItem.dcf_value" class="result-row">
              <span class="result-label">DCF估值</span>
              <span class="result-value">{{ formatMoney(selectedItem.dcf_value) }}</span>
            </div>
            <div v-if="selectedItem.pe_value" class="result-row">
              <span class="result-label">P/E估值</span>
              <span class="result-value">{{ formatMoney(selectedItem.pe_value) }}</span>
            </div>
            <div v-if="selectedItem.ps_value" class="result-row">
              <span class="result-label">P/S估值</span>
              <span class="result-value">{{ formatMoney(selectedItem.ps_value) }}</span>
            </div>
            <div v-if="selectedItem.pb_value" class="result-row">
              <span class="result-label">P/B估值</span>
              <span class="result-value">{{ formatMoney(selectedItem.pb_value) }}</span>
            </div>
            <div v-if="selectedItem.ev_value" class="result-row">
              <span class="result-label">EV/EBITDA估值</span>
              <span class="result-value">{{ formatMoney(selectedItem.ev_value) }}</span>
            </div>
          </div>
        </div>
        <div class="modal-footer">
          <button @click="loadToResultPage" class="btn-primary" :disabled="!hasCompleteData">
            {{ hasCompleteData ? '加载到结果页' : '⚠️ 此记录无完整详情' }}
          </button>
          <button @click="closeModal" class="btn-secondary">关闭</button>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, computed } from 'vue'
import { useRouter } from 'vue-router'
import axios from 'axios'

const router = useRouter()
const history = ref<any[]>([])
const loading = ref(false)
const selectedItem = ref<any>(null)
const showModal = ref(false)

// 计算选中项是否有完整数据
const hasCompleteData = computed(() => {
  return selectedItem.value?.results !== undefined && selectedItem.value?.results !== null
})

onMounted(() => {
  loadHistory()
})

const loadHistory = async () => {
  loading.value = true
  try {
    const response = await axios.get('http://localhost:8000/api/history?limit=50')
    if (response.data.success) {
      history.value = response.data.history
    }
  } catch (err: any) {
    console.error('加载历史记录失败:', err)
  } finally {
    loading.value = false
  }
}

const viewHistoryItem = async (id: number) => {
  try {
    const response = await axios.get(`http://localhost:8000/api/history/${id}`)
    if (response.data.success) {
      selectedItem.value = response.data.history
      showModal.value = true
    }
  } catch (err: any) {
    console.error('加载历史记录项失败:', err)
  }
}

const closeModal = () => {
  showModal.value = false
  selectedItem.value = null
}

const loadToResultPage = () => {
  if (!selectedItem.value) return

  // 检查是否有完整的 results 数据
  if (!selectedItem.value.results) {
    alert('此历史记录没有完整的估值详情数据。请重新进行估值分析以获取完整数据。')
    return
  }

  // 存储到sessionStorage并跳转到结果页
  sessionStorage.setItem('valuationResults', JSON.stringify(selectedItem.value))
  closeModal()
  router.push('/valuation/result')
}

const formatMoney = (value: number) => {
  if (!value) return '--'
  return (value).toFixed(2) + ' 亿元'
}

const formatDate = (dateStr: string) => {
  if (!dateStr) return '--'
  const date = new Date(dateStr)
  return date.toLocaleDateString('zh-CN', { year: 'numeric', month: '2-digit', day: '2-digit' })
}

const formatDateTime = (dateStr: string) => {
  if (!dateStr) return '--'
  const date = new Date(dateStr)
  return date.toLocaleString('zh-CN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  })
}
</script>

<style scoped>
.history-page {
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
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.card-actions {
  display: flex;
  justify-content: flex-end;
  margin-bottom: 20px;
}

.btn-refresh {
  background: #667eea;
  color: white;
  border: none;
  padding: 8px 16px;
  border-radius: 6px;
  cursor: pointer;
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
  padding: 60px 20px;
  color: #666;
}

.no-history p {
  margin-bottom: 10px;
}

.no-history .hint {
  font-size: 0.9em;
  color: #999;
  margin-bottom: 20px;
}

.btn-primary {
  display: inline-block;
  background: #667eea;
  color: white;
  text-decoration: none;
  padding: 10px 24px;
  border-radius: 6px;
  transition: all 0.3s;
}

.btn-primary:hover {
  background: #5568d3;
  transform: translateY(-2px);
}

.btn-secondary {
  background: #e0e0e0;
  color: #333;
  border: none;
  padding: 10px 24px;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.3s;
}

.btn-secondary:hover {
  background: #d0d0d0;
}

.history-list {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
  gap: 20px;
}

.history-item {
  background: #f8f9fa;
  padding: 20px;
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.3s;
  border: 2px solid transparent;
}

.history-item:hover {
  border-color: #667eea;
  box-shadow: 0 4px 12px rgba(102, 126, 234, 0.2);
  transform: translateY(-2px);
}

.history-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 10px;
}

.history-company {
  font-size: 1.1em;
  font-weight: 600;
  color: #333;
}

.history-date {
  font-size: 0.85em;
  color: #999;
}

.history-meta {
  display: flex;
  gap: 10px;
  margin-bottom: 12px;
}

.history-industry,
.history-stage {
  font-size: 0.85em;
  padding: 3px 10px;
  background: #e8f0ff;
  color: #667eea;
  border-radius: 12px;
}

.history-values {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.value-item {
  font-size: 0.9em;
  color: #555;
  background: white;
  padding: 6px 12px;
  border-radius: 4px;
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
  align-items: center;
  justify-content: center;
  z-index: 1000;
}

.modal-content {
  background: white;
  border-radius: 12px;
  max-width: 600px;
  width: 90%;
  max-height: 80vh;
  overflow-y: auto;
}

.modal-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 20px 25px;
  border-bottom: 1px solid #e0e0e0;
}

.modal-header h2 {
  margin: 0;
  font-size: 1.3em;
  color: #333;
}

.btn-close {
  background: none;
  border: none;
  font-size: 1.8em;
  cursor: pointer;
  color: #999;
  width: 32px;
  height: 32px;
  padding: 0;
  line-height: 1;
}

.btn-close:hover {
  color: #333;
}

.modal-body {
  padding: 25px;
}

.detail-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 15px;
  margin-bottom: 25px;
}

.detail-item {
  display: flex;
  flex-direction: column;
  gap: 5px;
}

.detail-label {
  font-size: 0.85em;
  color: #999;
}

.detail-value {
  font-size: 1em;
  color: #333;
  font-weight: 500;
}

.valuation-results {
  margin-top: 20px;
}

.valuation-results h3 {
  margin-bottom: 15px;
  font-size: 1.1em;
  color: #333;
}

.result-row {
  display: flex;
  justify-content: space-between;
  padding: 12px 15px;
  background: #f8f9fa;
  border-radius: 6px;
  margin-bottom: 8px;
}

.result-label {
  color: #666;
}

.result-value {
  font-weight: 600;
  color: #667eea;
}

.modal-footer {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
  padding: 15px 25px;
  border-top: 1px solid #e0e0e0;
}
</style>
