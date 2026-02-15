<template>
  <div class="valuation-input">
    <div class="header">
      <h1>📊 公司估值</h1>
      <p>输入公司基本信息和财务数据进行估值</p>
    </div>

    <div class="form-card">
      <div class="section-title">公司基本信息</div>
      <div class="form-grid">
        <div class="form-group">
          <label>公司名称</label>
          <input v-model="form.name" type="text" placeholder="输入公司名称" />
        </div>
        <div class="form-group">
          <label>所属行业(申万三级分类)</label>
          <div class="industry-cascade">
            <select v-model="selectedL1" @change="onL1Change" class="industry-select">
              <option value="">请选择一级行业...</option>
              <option v-for="l1 in shenwanIndustries" :key="l1.code" :value="l1.code">{{ l1.name }}</option>
            </select>
            <select v-model="selectedL2" @change="onL2Change" class="industry-select" :disabled="!selectedL1">
              <option value="">请选择二级行业...</option>
              <option v-for="l2 in l2Industries" :key="l2.code" :value="l2.code">{{ l2.name }}</option>
            </select>
            <select v-model="selectedL3" @change="onL3Change" class="industry-select" :disabled="!selectedL2">
              <option value="">请选择三级行业...</option>
              <option v-for="l3 in l3Industries" :key="l3.code" :value="l3.code">{{ l3.name }}</option>
            </select>
          </div>
          <input v-model="form.industry" type="hidden" />
          <div class="industry-selected" v-if="form.industry">
            已选择: {{ selectedIndustryPath }}
          </div>
        </div>
        <div class="form-group">
          <label>发展阶段</label>
          <select v-model="form.stage">
            <option value="早期">早期</option>
            <option value="成长期">成长期</option>
            <option value="成熟期">成熟期</option>
            <option value="上市公司">上市公司</option>
          </select>
        </div>
      </div>
    </div>

    <div class="form-card">
      <div class="section-title">财务数据(单位: 万元)</div>

      <!-- 上市公司Tushare导入区域 -->
      <div v-if="form.stage === '上市公司'" class="tushare-import-section">
        <div class="tushare-input-group">
          <label class="tushare-label">股票代码</label>
          <input
            v-model="stockCode"
            type="text"
            placeholder="例如: 000001.SZ (平安银行)"
            class="tushare-input"
            @keyup.enter="importStockFinancialData"
          />
          <button
            @click="importStockFinancialData"
            class="btn-tushare-import"
            type="button"
            :disabled="!stockCode || stockImporting">
            {{ stockImporting ? "导入中..." : "📥 从Tushare导入" }}
          </button>
        </div>
        <div class="tushare-hint">
          💡 提示:请输入6位数字股票代码+交易所后缀(如 .SZ 或 .SH)
        </div>
        <div v-if="stockImportError" class="stock-import-error">
          {{ stockImportError }}
          <div class="error-suggestions" v-if="stockImportError.includes('未找到')">
            <p>可能的原因:</p>
            <ul>
              <li>股票代码不存在或已退市</li>
              <li>该股票在Tushare数据库中暂无数据</li>
              <li>股票代码格式不正确(应为6位数字+.SZ/.SH)</li>
            </ul>
            <p>建议:尝试使用知名的蓝筹股,如 000001.SZ(平安银行)、000002.SZ(万科A)等</p>
          </div>
        </div>
        <div v-if="stockImportSuccess" class="stock-import-success">
          ✓ 财务数据已成功导入
        </div>
      </div>

      <div class="form-grid">
        <div class="form-group">
          <label>营业收入</label>
          <input v-model.number="form.revenue" type="number" placeholder="50000" />
        </div>
        <div class="form-group">
          <label>净利润</label>
          <input v-model.number="form.net_income" type="number" placeholder="8000" />
        </div>
        <div class="form-group">
          <label>净资产</label>
          <input v-model.number="form.net_assets" type="number" placeholder="20000" />
        </div>
        <div class="form-group">
          <label>EBITDA</label>
          <input v-model.number="form.ebitda" type="number" placeholder="12000" />
        </div>
        <div class="form-group">
          <label>总债务</label>
          <input v-model.number="form.total_debt" type="number" placeholder="5000" />
        </div>
        <div class="form-group">
          <label>货币资金</label>
          <input v-model.number="form.cash_and_equivalents" type="number" placeholder="2000" />
        </div>
      </div>
    </div>

    <div class="form-card">
      <div class="section-title">预测参数</div>
      <div class="form-grid">
        <div class="form-group">
          <label>预期增长率 (%)</label>
          <input v-model.number="form.growth_rate" type="number" step="0.1" placeholder="25" />
        </div>
        <div class="form-group">
          <label>营业利润率 (%)</label>
          <input v-model.number="form.operating_margin" type="number" step="0.1" placeholder="25" />
        </div>
        <div class="form-group">
          <label>贝塔系数 (β)</label>
          <input v-model.number="form.beta" type="number" step="0.1" placeholder="1.2" />
        </div>
        <div class="form-group">
          <label>无风险利率</label>
          <input v-model.number="form.risk_free_rate" type="number" step="0.01" placeholder="0.03" />
        </div>
        <div class="form-group">
          <label>市场风险溢价</label>
          <input v-model.number="form.market_risk_premium" type="number" step="0.01" placeholder="0.07" />
        </div>
        <div class="form-group">
          <label>永续增长率</label>
          <input v-model.number="form.terminal_growth_rate" type="number" step="0.005" placeholder="0.025" />
        </div>
      </div>
    </div>

    <!-- 可比公司数据输入 -->
    <div class="form-card">
      <div class="section-title">
        可比公司数据(可选,用于相对估值)
      </div>

      <!-- 导入选项 -->
      <div class="import-options">
        <button @click="openImportModal" class="btn-import" type="button">
          📥 从Tushare导入行业上市公司
        </button>
        <button @click="addComparable" class="btn-add" type="button">
          ✏️ 手动添加可比公司
        </button>
        <button @click="addSampleComparables" class="btn-secondary" type="button">
          📋 使用示例数据
        </button>
      </div>

      <!-- 已选可比公司列表 -->
      <div v-if="comparables.length > 0" class="comparables-header">
        <h3>已选择 {{ comparables.length }} 家可比公司</h3>
        <button @click="clearComparables" class="btn-clear" type="button">清空</button>
      </div>

      <div v-if="comparables.length === 0" class="no-comparables">
        <p>暂无可比公司数据,将仅使用DCF估值</p>
        <p class="hint">建议从Tushare导入目标公司所在行业的上市公司作为可比公司</p>
      </div>

      <div v-else class="comparables-list">
        <div v-for="(comp, idx) in comparables" :key="idx" class="comparable-card">
          <div class="comparable-header">
            <span class="comp-name">{{ comp.name }}</span>
            <span class="comp-info">{{ comp.industry }} | 收入: {{ (comp.revenue/10000).toFixed(1) }}亿 | 净利: {{ (comp.net_income/10000).toFixed(1) }}亿</span>
            <button @click="removeComparable(idx)" class="btn-remove" type="button">删除</button>
          </div>
          <div class="form-grid">
            <div class="form-group">
              <label>P/E倍数</label>
              <input v-model.number="comp.pe_ratio" type="number" step="0.1" />
            </div>
            <div class="form-group">
              <label>P/S倍数</label>
              <input v-model.number="comp.ps_ratio" type="number" step="0.1" />
            </div>
            <div class="form-group">
              <label>P/B倍数</label>
              <input v-model.number="comp.pb_ratio" type="number" step="0.1" />
            </div>
            <div class="form-group">
              <label>EV/EBITDA倍数</label>
              <input v-model.number="comp.ev_ebitda" type="number" step="0.1" />
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Tushare导入弹窗 -->
    <div v-if="showImportModal" class="modal-overlay" @click.self="showImportModal = false">
      <div class="modal-content">
        <div class="modal-header">
          <h2>从Tushare导入可比公司</h2>
          <button @click="showImportModal = false" class="btn-close">×</button>
        </div>
        <div class="modal-body">
          <div class="form-group">
            <label>选择行业(申万三级分类)</label>
            <div class="industry-cascade">
              <select v-model="importSelectedL1" @change="onImportL1Change" class="industry-select">
                <option value="">请选择一级行业...</option>
                <option v-for="l1 in shenwanIndustries" :key="l1.code" :value="l1.code">{{ l1.name }}</option>
              </select>
              <select v-model="importSelectedL2" @change="onImportL2Change" class="industry-select" :disabled="!importSelectedL1">
                <option value="">请选择二级行业...</option>
                <option v-for="l2 in importL2Industries" :key="l2.code" :value="l2.code">{{ l2.name }}</option>
              </select>
              <select v-model="importSelectedL3" @change="onImportL3Change" class="industry-select" :disabled="!importSelectedL2">
                <option value="">请选择三级行业...</option>
                <option v-for="l3 in importL3Industries" :key="l3.code" :value="l3.code">{{ l3.name }}</option>
              </select>
            </div>
            <input v-model="importIndustry" type="hidden" />
            <div class="industry-selected" v-if="importIndustry">
              已选择: {{ importSelectedIndustryPath }}
            </div>
          </div>

          <div class="form-group">
            <label>筛选条件</label>
            <div class="filter-grid">
              <div>
                <label>最小市值(亿元)</label>
                <input v-model.number="importMinMarketCap" type="number" placeholder="不限制">
              </div>
              <div>
                <label>最大市值(亿元)</label>
                <input v-model.number="importMaxMarketCap" type="number" placeholder="不限制">
              </div>
              <div>
                <label>返回数量</label>
                <input v-model.number="importLimit" type="number" min="5" max="50" value="20">
              </div>
            </div>
          </div>

          <button @click="importFromTushare" class="btn-primary btn-block" :disabled="importing">
            {{ importing ? '导入中...' : '🔍 获取公司列表' }}
          </button>

          <div v-if="importError" class="import-error">
            {{ importError }}
          </div>

          <!-- 导入的公司列表 -->
          <div v-if="availableCompanies.length > 0" class="companies-selection-with-actions">
            <div class="selection-header">
              <h3>找到 {{ availableCompanies.length }} 家上市公司</h3>
              <p class="hint">请根据业务相似度和规模选择合适的可比公司</p>
              <div class="selection-actions-top">
                <button @click="selectAllCompanies" class="btn-small">全选</button>
                <button @click="clearSelection" class="btn-small">清空选择</button>
              </div>
            </div>

            <div class="companies-list-with-footer">
              <div class="companies-list">
                <div v-for="company in availableCompanies" :key="company.ts_code"
                     :class="['company-item', { selected: isCompanySelected(company.ts_code) }]"
                     @click="toggleCompanySelection(company.ts_code)">
                  <div class="company-checkbox">
                    <input type="checkbox" :checked="isCompanySelected(company.ts_code)" readonly>
                  </div>
                  <div class="company-info">
                    <div class="company-name">{{ company.name }}</div>
                    <div class="company-details">
                      <span>代码: {{ company.ts_code }}</span>
                      <span>收入: {{ (company.revenue/10000).toFixed(1) }}亿</span>
                      <span>净利: {{ (company.net_income/10000).toFixed(1) }}亿</span>
                      <span v-if="company.pe_ratio">P/E: {{ company.pe_ratio?.toFixed(1) }}</span>
                      <span v-if="company.pb_ratio">P/B: {{ company.pb_ratio?.toFixed(1) }}</span>
                    </div>
                  </div>
                </div>
              </div>

              <!-- 固定在底部的确认按钮 -->
              <div class="companies-footer">
                <div class="footer-summary">
                  已选择 <strong>{{ selectedCompaniesCount }}</strong> 家公司
                </div>
                <button @click="addSelectedCompanies" class="btn-confirm-add" :disabled="selectedCompaniesCount === 0">
                  ✓ 确认添加选中的公司
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div class="actions">
      <button class="btn btn-primary" @click="startValuation" :disabled="loading" onclick="console.log('原生点击事件触发!')">
        {{ loading ? '计算中...' : '🚀 开始估值' }}
      </button>
      <button class="btn btn-secondary" @click="resetForm">🔄 重置</button>
    </div>

    <div v-if="error" class="error">{{ error }}</div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick } from 'vue'
import { useRouter } from 'vue-router'
import { valuationAPI, scenarioAPI, stressTestAPI, sensitivityAPI, dataAPI } from '../services/api'
import axios from 'axios'

const router = useRouter()

const form = ref({
  name: "云数科技有限公司",
  industry: '软件服务',
  stage: '成长期',
  revenue: 50000,
  net_income: 8000,
  net_assets: 20000,
  ebitda: 12000,
  total_debt: 5000,
  cash_and_equivalents: 2000,
  growth_rate: 25,
  operating_margin: 25,
  beta: 1.2,
  risk_free_rate: 0.03,
  market_risk_premium: 0.07,
  terminal_growth_rate: 0.025
} as any)

const comparables = ref<any[]>([])
const loading = ref(false)
const error = ref('')

// 上市公司Tushare导入相关
const stockCode = ref('')
const stockImporting = ref(false)
const stockImportError = ref('')
const stockImportSuccess = ref(false)

// 申万三级分类级联选择
const shenwanIndustries = ref<any[]>([])
const selectedL1 = ref('')
const selectedL2 = ref('')
const selectedL3 = ref('')

// 加载申万行业分类数据
fetch('/shenwan_industries.json')
  .then(res => res.json())
  .then(data => {
    shenwanIndustries.value = data
    // 设置默认选择(计算机 -> IT服务 -> 垂直应用软件)
    const l1 = data.find((i: any) => i.name === '计算机')
    if (l1) {
      selectedL1.value = l1.code
      onL1Change()
      const l2 = l1.children?.find((c: any) => c.name === '软件开发')
      if (l2) {
        selectedL2.value = l2.code
        onL2Change()
        const l3 = l2.children?.find((c: any) => c.name === '垂直应用软件')
        if (l3) {
          selectedL3.value = l3.code
          onL3Change()
        }
      }
    }
  })
  .catch(err => console.error('加载行业分类失败:', err))

// 二级行业列表
const l2Industries = computed(() => {
  if (!selectedL1.value) return []
  const l1 = shenwanIndustries.value.find((i: any) => i.code === selectedL1.value)
  return l1?.children || []
})

// 三级行业列表
const l3Industries = computed(() => {
  if (!selectedL2.value) return []
  const l2 = l2Industries.value.find((i: any) => i.code === selectedL2.value)
  return l2?.children || []
})

// 显示选择的完整路径
const selectedIndustryPath = computed(() => {
  const l1 = shenwanIndustries.value.find((i: any) => i.code === selectedL1.value)
  const l2 = l2Industries.value.find((i: any) => i.code === selectedL2.value)
  const l3 = l3Industries.value.find((i: any) => i.code === selectedL3.value)
  const parts = [l1?.name, l2?.name, l3?.name].filter(Boolean)
  return parts.join(' > ')
})

// L1选择变化
const onL1Change = () => {
  selectedL2.value = ''
  selectedL3.value = ''
  form.value.industry = ''
}

// L2选择变化
const onL2Change = () => {
  selectedL3.value = ''
  form.value.industry = ''
}

// L3选择变化
const onL3Change = () => {
  if (selectedL3.value) {
    form.value.industry = selectedL3.value
    // 如果导入弹窗中的行业为空,自动填充
    if (!importIndustry.value) {
      importIndustry.value = selectedL3.value
    }
  } else {
    form.value.industry = ''
  }
}

// Tushare导入相关
const showImportModal = ref(false)
const importIndustry = ref('')
const importMinMarketCap = ref<number | null>(null)
const importMaxMarketCap = ref<number | null>(null)
const importLimit = ref(20)
const importing = ref(false)
const importError = ref('')
const availableCompanies = ref<any[]>([])
const selectedCompanyCodes = ref<Set<string>>(new Set())

const selectedCompaniesCount = computed(() => selectedCompanyCodes.value.size)

// 导入弹窗的申万三级分类级联选择
const importSelectedL1 = ref('')
const importSelectedL2 = ref('')
const importSelectedL3 = ref('')

// 导入弹窗的二级行业列表
const importL2Industries = computed(() => {
  if (!importSelectedL1.value) return []
  const l1 = shenwanIndustries.value.find((i: any) => i.code === importSelectedL1.value)
  return l1?.children || []
})

// 导入弹窗的三级行业列表
const importL3Industries = computed(() => {
  if (!importSelectedL2.value) return []
  const l2 = importL2Industries.value.find((i: any) => i.code === importSelectedL2.value)
  return l2?.children || []
})

// 导入弹窗的已选择完整路径
const importSelectedIndustryPath = computed(() => {
  const l1 = shenwanIndustries.value.find((i: any) => i.code === importSelectedL1.value)
  const l2 = importL2Industries.value.find((i: any) => i.code === importSelectedL2.value)
  const l3 = importL3Industries.value.find((i: any) => i.code === importSelectedL3.value)
  const parts = [l1?.name, l2?.name, l3?.name].filter(Boolean)
  return parts.join(' > ')
})

// 导入弹窗L1选择变化
const onImportL1Change = () => {
  importSelectedL2.value = ''
  importSelectedL3.value = ''
  importIndustry.value = ''
}

// 导入弹窗L2选择变化
const onImportL2Change = () => {
  importSelectedL3.value = ''
  importIndustry.value = ''
}

// 导入弹窗L3选择变化
const onImportL3Change = () => {
  if (importSelectedL3.value) {
    importIndustry.value = importSelectedL3.value
    // 当在导入弹窗选择行业时,如果目标公司没有行业,自动填充
    if (!form.value.industry) {
      form.value.industry = importSelectedL3.value
    }
  } else {
    importIndustry.value = ''
  }
}

const addComparable = () => {
  comparables.value.push({
    name: '',
    industry: '',
    revenue: 0,
    net_income: 0,
    net_assets: 0,
    ebitda: null,
    pe_ratio: null,
    ps_ratio: null,
    pb_ratio: null,
    ev_ebitda: null,
    growth_rate: null
  })
}

const removeComparable = (idx: number) => {
  comparables.value.splice(idx, 1)
}

const addSampleComparables = () => {
  comparables.value = [
    {
      name: '金山云',
      industry: '软件服务',
      revenue: 80000,
      net_income: 5000,
      net_assets: 35000,
      ebitda: 10000,
      pe_ratio: 30.0,
      ps_ratio: 5.5,
      pb_ratio: 3.8,
      ev_ebitda: 22.0,
      growth_rate: 0.30
    },
    {
      name: '用友网络',
      industry: '软件服务',
      revenue: 95000,
      net_income: 12000,
      net_assets: 45000,
      ebitda: 18000,
      pe_ratio: 45.0,
      ps_ratio: 6.8,
      pb_ratio: 5.2,
      ev_ebitda: 28.0,
      growth_rate: 0.20
    },
    {
      name: '恒生电子',
      industry: '软件服务',
      revenue: 70000,
      net_income: 15000,
      net_assets: 40000,
      ebitda: 20000,
      pe_ratio: 35.0,
      ps_ratio: 7.5,
      pb_ratio: 4.5,
      ev_ebitda: 25.0,
      growth_rate: 0.18
    }
  ]
}

// Tushare导入相关方法
const clearComparables = () => {
  comparables.value = []
}

// 从Tushare导入上市公司财务数据
const importStockFinancialData = async () => {
  if (!stockCode.value) {
    stockImportError.value = '请输入股票代码'
    return
  }

  // 清除之前的错误和成功状态
  stockImportError.value = ''
  stockImportSuccess.value = false
  stockImporting.value = true

  try {
    const response = await dataAPI.getStockData(stockCode.value)

    if (response.data && response.data.success) {
      const data = response.data.data

      // 填充财务数据到表单(注意:后端返回的单位是"元",需要转换为"万元",即除以10000)
      if (data.revenue !== undefined) form.value.revenue = Math.round(data.revenue / 10000)
      if (data.net_income !== undefined) form.value.net_income = Math.round(data.net_income / 10000)
      if (data.net_assets !== undefined) form.value.net_assets = Math.round(data.net_assets / 10000)
      if (data.ebitda !== undefined) form.value.ebitda = Math.round(data.ebitda / 10000)
      if (data.total_debt !== undefined) form.value.total_debt = Math.round(data.total_debt / 10000)
      if (data.cash_and_equivalents !== undefined) form.value.cash_and_equivalents = Math.round(data.cash_and_equivalents / 10000)

      // 如果API返回了公司名称,更新表单
      if (data.name) form.value.name = data.name

      stockImportSuccess.value = true

      // 3秒后清除成功提示
      setTimeout(() => {
        stockImportSuccess.value = false
      }, 3000)
    } else {
      stockImportError.value = "未找到该股票的财务数据,请检查股票代码是否正确"
    }
  } catch (err: any) {
    console.error('导入财务数据失败:', err)
    if (err.response?.status === 404) {
      stockImportError.value = `未找到股票代码 "${stockCode.value}" 的数据。请确认股票代码是否正确,或尝试其他股票代码。`
    } else if (err.response?.data?.detail) {
      stockImportError.value = err.response.data.detail
    } else {
      stockImportError.value = `导入失败:${err.message || '未知错误'}。请检查网络连接。`
    }
  } finally {
    stockImporting.value = false
  }
}

// 打开导入弹窗时,同步主表单的行业选择
const openImportModal = () => {
  showImportModal.value = true
  // 如果主表单已经选择了行业,同步到导入弹窗
  if (selectedL1.value) {
    importSelectedL1.value = selectedL1.value
    if (selectedL2.value) {
      importSelectedL2.value = selectedL2.value
      if (selectedL3.value) {
        importSelectedL3.value = selectedL3.value
        importIndustry.value = selectedL3.value
      }
    }
  }
}

const importFromTushare = async () => {
  if (!importIndustry.value) {
    importError.value = "请先选择行业"
    return
  }

  importing.value = true
  importError.value = ''

  try {
    const params: any = {
      limit: importLimit.value
    }

    if (importMinMarketCap.value) params.market_cap_min = importMinMarketCap.value
    if (importMaxMarketCap.value) params.market_cap_max = importMaxMarketCap.value

    // 使用URL编码处理中文行业名
    const encodedIndustry = encodeURIComponent(importIndustry.value)
    const response = await axios.get(
      `http://localhost:8000/api/data/comparable/${encodedIndustry}`,
      { params }
    )

    if (response.data.success) {
      availableCompanies.value = response.data.companies
      selectedCompanyCodes.value.clear()

      if (availableCompanies.value.length === 0) {
        importError.value = '未找到符合条件的公司,请尝试调整筛选条件或选择其他行业'
      }
    }
  } catch (err: any) {
    console.error("导入失败:", err)
    importError.value = "导入失败: " + (err.response?.data?.detail || err.message)
  } finally {
    importing.value = false
  }
}

const isCompanySelected = (tsCode: string) => {
  return selectedCompanyCodes.value.has(tsCode)
}

const toggleCompanySelection = (tsCode: string) => {
  if (selectedCompanyCodes.value.has(tsCode)) {
    selectedCompanyCodes.value.delete(tsCode)
  } else {
    selectedCompanyCodes.value.add(tsCode)
  }
}

const selectAllCompanies = () => {
  availableCompanies.value.forEach(c => {
    selectedCompanyCodes.value.add(c.ts_code)
  })
}

const clearSelection = () => {
  selectedCompanyCodes.value.clear()
}

const addSelectedCompanies = () => {
  selectedCompanyCodes.value.forEach(tsCode => {
    const company = availableCompanies.value.find(c => c.ts_code === tsCode)
    if (company && !comparables.value.some(c => c.ts_code === tsCode)) {
      comparables.value.push({
        name: company.name,
        ts_code: company.ts_code,
        industry: company.industry,
        revenue: company.revenue,
        net_income: company.net_income,
        net_assets: company.net_assets,
        ebitda: company.ebitda,
        pe_ratio: company.pe_ratio,
        ps_ratio: company.ps_ratio,
        pb_ratio: company.pb_ratio,
        ev_ebitda: company.ev_ebitda,
        growth_rate: company.growth_rate
      })
    }
  })

  showImportModal.value = false
  selectedCompanyCodes.value.clear()
  availableCompanies.value = []
}

const getErrorMessage = (err: any): string => {
  console.error("详细错误:", err)

  if (err.response?.data) {
    const data = err.response.data
    if (typeof data === 'string') {
      return data
    }
    if (data.detail) {
      if (typeof data.detail === 'string') {
        return data.detail
      }
      return JSON.stringify(data.detail)
    }
    return JSON.stringify(data)
  }

  if (err.message) {
    return err.message
  }

  return JSON.stringify(err)
}

const startValuation = async () => {
  console.log('=== 开始估值按钮被点击 ===')
  console.log('当前表单数据:', JSON.parse(JSON.stringify(form.value)))

  error.value = ''
  loading.value = true

  console.log('loading.value已设为true,按钮应该显示"计算中..."')

  try {
    const company = {
      ...form.value,
      growth_rate: form.value.growth_rate / 100,
      operating_margin: form.value.operating_margin / 100
    }

    console.log('公司数据:', company)
    console.log('可比公司数量:', comparables.value.length)

    // 并行执行多个估值请求
    const requests = [
      valuationAPI.dcf(company),
      scenarioAPI.analyze(company),
      stressTestAPI.full(company),
      sensitivityAPI.comprehensive(company)
    ]

    // 如果有可比公司,添加相对估值
    if (comparables.value.length > 0) {
      console.log('可比公司数据:', comparables.value)
      requests.unshift(valuationAPI.relative(company, comparables.value))
    }

    console.log('开始并行请求API，共', requests.length, '个请求')
    const results = await Promise.all(requests)
    console.log('所有API请求已完成')

    let dcfResult, scenarioResult, stressResult, sensitivityResult, relativeResult

    if (comparables.value.length > 0) {
      [relativeResult, dcfResult, scenarioResult, stressResult, sensitivityResult] = results
    } else {
      [dcfResult, scenarioResult, stressResult, sensitivityResult] = results
    }

    console.log('API响应结果:', {
      relative: relativeResult,
      dcf: dcfResult,
      dcfData: dcfResult?.data,
      dcfResult: dcfResult?.data?.result,
      scenario: scenarioResult
    })

    // 检查API响应状态
    if (!dcfResult?.data?.success) {
      throw new Error('DCF估值失败')
    }

    // 存储结果到sessionStorage用于结果页展示
    const resultsToStore = {
      relative: relativeResult?.data,
      dcf: dcfResult?.data,  // dcfResult.data = {success: true, result: {...}}
      scenario: scenarioResult?.data,
      stress: stressResult?.data,
      sensitivity: sensitivityResult?.data,
      company: form.value,
      comparables: comparables.value
    }
    console.log('准备存储到sessionStorage的数据:', resultsToStore)
    console.log('DCF数据详情:', resultsToStore.dcf)

    // 确保sessionStorage保存完成后再跳转
    try {
      console.log('开始序列化数据...')
      const jsonStr = JSON.stringify(resultsToStore)
      console.log('序列化后的JSON字符串长度:', jsonStr.length)
      console.log('JSON字符串预览(前200字符):', jsonStr.substring(0, 200))

      console.log('开始保存到sessionStorage...')
      sessionStorage.setItem('valuationResults', jsonStr)
      console.log('✅ sessionStorage.setItem调用成功')

      // 立即验证
      const stored = sessionStorage.getItem('valuationResults')
      console.log('验证存储 - 立即读取结果:', stored ? '成功' : 'NULL!')
      if (stored) {
        try {
          const parsed = JSON.parse(stored)
          console.log('验证存储 - 解析成功, 数据键:', Object.keys(parsed))
          console.log('验证存储 - DCF结果:', parsed.dcf)
          console.log('验证存储 - 相对估值结果:', parsed.relative)
        } catch (parseErr) {
          console.error('验证存储 - JSON解析失败:', parseErr)
        }
      } else {
        console.error('验证存储 - 读取失败,数据未保存!')
        throw new Error('sessionStorage数据保存失败，无法读取已保存的数据')
      }

      // 小延迟确保存储完成
      await nextTick()
      console.log('即将跳转到结果页...')
      router.push('/valuation/result')
      console.log('✅ router.push调用完成')
    } catch (err: unknown) {
      console.error('sessionStorage操作失败:', err)
      if (err instanceof Error) {
        console.error('错误堆栈:', err.stack)
        error.value = '数据保存失败:' + err.message
      } else {
        error.value = '数据保存失败:未知错误'
      }
      loading.value = false
      return
    }
  } catch (err: any) {
    console.error('startValuation发生错误:', err)
    error.value = '估值计算失败: ' + getErrorMessage(err)
    loading.value = false
  } finally {
    if (loading.value) {
      loading.value = false
    }
  }
}

const resetForm = () => {
  form.value = {
    name: '',
    industry: '',
    stage: '成长期',
    revenue: 0,
    net_income: 0,
    net_assets: 0,
    ebitda: 0,
    total_debt: 0,
    cash_and_equivalents: 0,
    growth_rate: 15,
    operating_margin: 20,
    beta: 1.0,
    risk_free_rate: 0.03,
    market_risk_premium: 0.07,
    terminal_growth_rate: 0.025
  }
  comparables.value = []
}
</script>

<style scoped>
.valuation-input {
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

.form-card {
  background: white;
  padding: 25px;
  border-radius: 12px;
  margin-bottom: 20px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.section-title {
  font-size: 1.2em;
  font-weight: bold;
  color: #667eea;
  margin-bottom: 20px;
  padding-bottom: 10px;
  border-bottom: 2px solid #667eea;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.btn-add {
  background: #667eea;
  color: white;
  border: none;
  padding: 6px 12px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 0.85em;
}

.btn-add:hover {
  background: #5568d3;
}

.form-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
  gap: 15px;
}

.form-group {
  display: flex;
  flex-direction: column;
}

.form-group label {
  margin-bottom: 5px;
  color: #555;
  font-weight: 500;
  font-size: 0.9em;
}

.form-group input,
.form-group select {
  padding: 10px;
  border: 1px solid #ddd;
  border-radius: 6px;
  font-size: 14px;
}

.form-group input:focus,
.form-group select:focus {
  outline: none;
  border-color: #667eea;
}

/* 上市公司Tushare导入区域样式 */
.tushare-import-section {
  background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
  border: 1px solid #667eea;
  border-radius: 8px;
  padding: 20px;
  margin-bottom: 20px;
}

.tushare-input-group {
  display: flex;
  gap: 15px;
  align-items: flex-end;
  flex-wrap: wrap;
}

.tushare-input-wrapper {
  display: flex;
  flex-direction: column;
  flex: 1;
  min-width: 250px;
}

.tushare-label {
  margin-bottom: 8px;
  font-weight: 600;
  color: #333;
  font-size: 0.95em;
}

.tushare-input {
  padding: 12px 16px;
  border: 2px solid #667eea;
  border-radius: 6px;
  font-size: 0.95em;
  background: white;
  transition: all 0.2s;
}

.tushare-input:focus {
  outline: none;
  border-color: #5568d3;
  box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
}

.btn-tushare-import {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  border: none;
  padding: 12px 24px;
  border-radius: 6px;
  cursor: pointer;
  font-size: 0.95em;
  font-weight: 600;
  white-space: nowrap;
  transition: all 0.2s;
}

.btn-tushare-import:hover:not(:disabled) {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
}

.btn-tushare-import:disabled {
  background: #ccc;
  cursor: not-allowed;
  opacity: 0.6;
}

.stock-import-error {
  margin-top: 15px;
  padding: 12px;
  background: #fee;
  border: 1px solid #f5c6cb;
  border-radius: 6px;
  color: #c62828;
  font-size: 0.9em;
}

.stock-import-success {
  margin-top: 15px;
  padding: 12px;
  background: #d4edda;
  border: 1px solid #c3e6cb;
  border-radius: 6px;
  color: #0f5132;
  font-size: 0.9em;
  font-weight: 500;
}

.tushare-hint {
  margin-top: 10px;
  padding: 10px 12px;
  background: #fffbeb;
  border-left: 3px solid #ffa500;
  border-radius: 4px;
  color: #666;
  font-size: 0.85em;
  line-height: 1.5;
}

.error-suggestions {
  margin-top: 12px;
  padding-top: 12px;
  border-top: 1px dashed #f5c6cb;
}

.error-suggestions p {
  margin: 8px 0;
  color: #555;
}

.error-suggestions ul {
  margin: 8px 0;
  padding-left: 20px;
  color: #666;
}

.error-suggestions li {
  margin: 4px 0;
}


.no-comparables {
  text-align: center;
  padding: 30px;
  background: #f8f9fa;
  border-radius: 8px;
}

.no-comparables p {
  color: #666;
  margin-bottom: 15px;
}

.comparables-list {
  display: flex;
  flex-direction: column;
  gap: 15px;
}

.comparable-card {
  background: #f8f9fa;
  padding: 20px;
  border-radius: 8px;
  border: 1px solid #e0e0e0;
}

.comparable-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 15px;
  font-weight: bold;
  color: #333;
}

.btn-remove {
  background: #ee6666;
  color: white;
  border: none;
  padding: 4px 10px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 0.85em;
}

.btn-remove:hover {
  background: #d65555;
}

.actions {
  display: flex;
  gap: 10px;
  justify-content: center;
  margin-top: 30px;
}

.btn {
  padding: 12px 40px;
  border: none;
  border-radius: 6px;
  font-size: 16px;
  cursor: pointer;
  transition: all 0.3s;
  font-weight: 500;
}

.btn-primary {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
}

.btn-primary:hover:not(:disabled) {
  transform: translateY(-2px);
  box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
}

.btn-primary:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.btn-secondary {
  background: #f0f0f0;
  color: #333;
  border: none;
  padding: 8px 16px;
  border-radius: 4px;
  cursor: pointer;
}

.btn-secondary:hover {
  background: #e0e0e0;
}

.error {
  background: #fee;
  color: #c33;
  padding: 15px;
  border-radius: 8px;
  margin-top: 20px;
  text-align: center;
}

/* 导入选项 */
.import-options {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  margin-bottom: 20px;
}

.btn-import {
  background: #764ba2;
  color: white;
  border: none;
  padding: 10px 20px;
  border-radius: 6px;
  cursor: pointer;
  font-size: 14px;
  transition: all 0.3s;
}

.btn-import:hover {
  background: #663a99;
  transform: translateY(-2px);
}

.comparables-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 15px;
  padding: 10px 15px;
  background: #f0f7ff;
  border-radius: 8px;
}

.comparables-header h3 {
  margin: 0;
  color: #333;
  font-size: 1.1em;
}

.btn-clear {
  background: #ee6666;
  color: white;
  border: none;
  padding: 6px 12px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 12px;
}

.btn-clear:hover {
  background: #d65555;
}

.hint {
  color: #999;
  font-size: 0.9em;
  margin-top: 5px;
}

/* 模态框 */
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
  width: 90%;
  max-width: 800px;
  max-height: 90vh;
  overflow-y: auto;
  box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
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
  font-size: 2em;
  color: #999;
  cursor: pointer;
  padding: 0;
  width: 30px;
  height: 30px;
  line-height: 1;
}

.btn-close:hover {
  color: #333;
}

.modal-body {
  padding: 25px;
}

.filter-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
  gap: 10px;
  margin-top: 10px;
}

.filter-grid > div {
  display: flex;
  flex-direction: column;
}

.filter-grid label {
  font-size: 0.85em;
  color: #666;
  margin-bottom: 5px;
}

.filter-grid input {
  padding: 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 14px;
}

.btn-block {
  width: 100%;
  margin-top: 15px;
}

.import-error {
  background: #fee;
  color: #c33;
  padding: 12px;
  border-radius: 6px;
  margin-top: 15px;
  text-align: center;
  font-size: 0.9em;
}

/* 公司选择列表 */
.companies-selection-with-actions {
  margin-top: 20px;
  border-top: 1px solid #e0e0e0;
  padding-top: 20px;
}

.selection-header {
  margin-bottom: 15px;
}

.selection-header h3 {
  margin: 0 0 5px 0;
  font-size: 1.1em;
  color: #333;
}

.selection-actions {
  display: flex;
  gap: 8px;
  margin-top: 10px;
  flex-wrap: wrap;
}

.btn-small {
  padding: 6px 12px;
  font-size: 13px;
  border-radius: 4px;
  border: 1px solid #ddd;
  background: white;
  cursor: pointer;
}

.btn-small:hover {
  background: #f5f5f5;
}

.btn-small.btn-primary {
  background: #667eea;
  color: white;
  border-color: #667eea;
}

.btn-small.btn-primary:hover {
  background: #5568d3;
}

.companies-list-with-footer {
  display: flex;
  flex-direction: column;
  max-height: 450px;
}

.companies-list {
  flex: 1;
  overflow-y: auto;
  border: 1px solid #e0e0e0;
  border-radius: 8px 8px 0 0;
  border-bottom: none;
  max-height: 380px;
}

/* 固定底部操作栏 */
.companies-footer {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 15px 20px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  border-radius: 0 0 8px 8px;
  box-shadow: 0 -2px 10px rgba(0, 0, 0, 0.1);
  position: sticky;
  bottom: 0;
}

.footer-summary {
  color: white;
  font-size: 14px;
}

.footer-summary strong {
  font-size: 18px;
  color: #91cc75;
}

.btn-confirm-add {
  background: white;
  color: #667eea;
  border: none;
  padding: 10px 24px;
  border-radius: 6px;
  font-size: 15px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.3s;
  display: flex;
  align-items: center;
  gap: 5px;
}

.btn-confirm-add:hover:not(:disabled) {
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
}

.btn-confirm-add:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.selection-actions-top {
  display: flex;
  gap: 8px;
  margin-top: 10px;
  flex-wrap: wrap;
}

.company-item {
  display: flex;
  align-items: center;
  padding: 12px 15px;
  border-bottom: 1px solid #f0f0f0;
  cursor: pointer;
  transition: background 0.2s;
}

.company-item:hover {
  background: #f8f9fa;
}

.company-item.selected {
  background: #e8f4ff;
}

.company-item:last-child {
  border-bottom: none;
}

.company-checkbox {
  margin-right: 12px;
}

.company-checkbox input {
  cursor: pointer;
}

.company-info {
  flex: 1;
}

.company-name {
  font-weight: 600;
  color: #333;
  margin-bottom: 5px;
}

.company-details {
  font-size: 0.85em;
  color: #666;
  display: flex;
  gap: 15px;
  flex-wrap: wrap;
}

.company-details span {
  white-space: nowrap;
}

/* 可比公司卡片更新 */
.comp-name {
  font-weight: 600;
  color: #333;
}

.comp-info {
  font-size: 0.9em;
  color: #666;
  flex: 1;
}

/* 申万三级分类级联选择器 */
.industry-cascade {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.industry-select {
  flex: 1 1 120px;
  min-width: 120px;
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 14px;
  background: white;
  cursor: pointer;
  transition: border-color 0.3s;
}

.industry-select:hover:not(:disabled) {
  border-color: #667eea;
}

.industry-select:disabled {
  background: #f5f5f5;
  color: #999;
  cursor: not-allowed;
}

.industry-selected {
  margin-top: 8px;
  padding: 8px 12px;
  background: #f0f7ff;
  border-radius: 4px;
  color: #667eea;
  font-size: 0.9em;
}
</style>
