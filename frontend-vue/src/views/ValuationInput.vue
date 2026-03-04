<template>
  <div class="valuation-input">
    <div class="header">
      <h1>📊 公司估值</h1>
      <p>输入公司基本信息和财务数据进行估值</p>
    </div>

    <div class="form-card">
      <div class="section-title">公司基本信息</div>
      <div class="form-grid company-info-grid">
        <div class="form-group">
          <label>公司名称</label>
          <input v-model="form.name" type="text" placeholder="输入公司名称" />
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
        <div class="form-group form-group-full-width">
          <label>所属行业(申万三级分类)</label>
          <div class="industry-cascade">
            <select v-model="selectedL1" @change="onL1Change" class="industry-select industry-select-wide">
              <option value="">请选择一级行业...</option>
              <option v-for="l1 in shenwanIndustries" :key="l1.code" :value="l1.code">{{ l1.name }}</option>
            </select>
            <select v-model="selectedL2" @change="onL2Change" class="industry-select industry-select-wide" :disabled="!selectedL1">
              <option value="">请选择二级行业...</option>
              <option v-for="l2 in l2Industries" :key="l2.code" :value="l2.code">{{ l2.name }}</option>
            </select>
            <input v-model="form.industry" type="hidden" />
          </div>
          <div class="industry-cascade industry-l3-row">
            <select v-model="selectedL3" @change="onL3Change" class="industry-select industry-select-wide" :disabled="!selectedL2">
              <option value="">请选择三级行业...</option>
              <option v-for="l3 in l3Industries" :key="l3.code" :value="l3.code">{{ l3.name }}</option>
            </select>
          </div>
          <div class="industry-selected" v-if="form.industry">
            已选择: {{ selectedIndustryPath }}
          </div>
        </div>
      </div>
    </div>

    <!-- 估值模式切换 -->
    <div class="form-card">
      <div class="section-title">估值模式</div>
      <div class="mode-switch">
        <button
          :class="['mode-btn', valuationMode === 'single' ? 'active' : '']"
          @click="valuationMode = 'single'"
          type="button">
          单产品估值
        </button>
        <button
          :class="['mode-btn', valuationMode === 'multi' ? 'active' : '']"
          @click="valuationMode = 'multi'"
          type="button">
          多产品估值
        </button>
      </div>
      <div class="mode-description">
        <template v-if="valuationMode === 'single'">
          单产品模式：使用统一的增长率、利润率等参数对公司整体估值
        </template>
        <template v-else>
          多产品模式：支持对多个产品/业务线分别估值，反映不同业务的差异
        </template>
      </div>
    </div>

    <!-- 多产品输入界面 -->
    <div v-if="valuationMode === 'multi'" class="form-card">
      <div class="section-title">
        产品/业务线配置
        <span class="weight-indicator" :class="{ 'error': Math.abs(totalWeight - 1) > 0.01 }">
          权重总和: {{ (totalWeight * 100).toFixed(1) }}%
        </span>
      </div>

      <div class="products-header">
        <button @click="addProduct" class="btn-add" type="button">
          ➕ 添加产品
        </button>
        <span class="products-count">共 {{ products.length }} 个产品</span>
        <button @click="autoCalculateWeights" class="btn-auto-calc" type="button" title="根据当前收入自动计算占比">
          🔄 自动计算占比
        </button>
      </div>

      <div class="products-list">
        <div v-for="(product, index) in products" :key="index" class="product-card">
          <div class="product-header">
            <div class="product-title">产品 {{ index + 1 }}</div>
            <button
              v-if="products.length > 1"
              @click="removeProduct(index)"
              class="btn-remove-product"
              type="button">
              ✕ 删除
            </button>
          </div>

          <div class="form-grid">
            <div class="form-group">
              <label>产品名称 *</label>
              <input v-model="product.name" type="text" placeholder="例如：云服务" />
            </div>
            <div class="form-group">
              <label>产品描述</label>
              <input v-model="product.description" type="text" placeholder="可选" />
            </div>
            <div class="form-group">
              <label>当前收入(万元) *</label>
              <input
                v-model.number="product.current_revenue"
                type="number"
                placeholder="50000"
                min="0"
                @input="autoCalculateWeights"
              />
            </div>
            <div class="form-group">
              <label>收入占比(%) *</label>
              <input
                v-model.number="product.revenue_weight"
                type="number"
                placeholder="自动计算"
                min="0"
                max="100"
                step="1"
                @input="manualWeightChange"
              />
              <small style="color: #999; font-size: 0.85em; margin-top: 4px; display: block;">
                {{ (product.revenue_weight * 100).toFixed(1) }}%
              </small>
            </div>
          </div>

          <div class="product-section">
            <div class="product-section-title">增长率（未来5年）</div>
            <div class="growth-years-input">
              <div v-for="(year, idx) in 5" :key="idx" class="year-input">
                <label>第{{ year }}年</label>
                <input
                  v-model.number="product.growth_rate_years[idx]"
                  type="number"
                  placeholder="15"
                  min="-50"
                  max="100"
                  step="1" />
                <span>%</span>
              </div>
            </div>
          </div>

          <div class="form-grid">
            <div class="form-group">
              <label>永续增长率(%)</label>
              <input v-model.number="product.terminal_growth_rate" type="number" step="0.5" placeholder="2.5" />
            </div>
            <div class="form-group">
              <label>毛利率(%)</label>
              <input v-model.number="product.gross_margin" type="number" placeholder="65" min="0" max="100" />
            </div>
            <div class="form-group">
              <label>营业利润率(%)</label>
              <input v-model.number="product.operating_margin" type="number" placeholder="25" min="0" max="100" />
            </div>
            <div class="form-group">
              <label>β系数（可选）</label>
              <input v-model.number="product.beta" type="number" step="0.1" placeholder="留空使用公司整体β" />
            </div>
            <div class="form-group">
              <label>资本支出/收入(%)</label>
              <input v-model.number="product.capex_ratio" type="number" placeholder="5" min="0" max="100" />
            </div>
            <div class="form-group">
              <label>营运资金变化/收入(%)</label>
              <input v-model.number="product.wc_change_ratio" type="number" placeholder="2" min="0" max="100" />
            </div>
          </div>
        </div>
      </div>
    </div>

    <div v-if="valuationMode === 'single'" class="form-card">
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
      <div class="section-title">预测参数 - 自由现金流</div>
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
          <label>资本支出/收入 (%)</label>
          <input v-model.number="form.capex_ratio" type="number" step="0.01" min="0" max="1" placeholder="0.05" />
          <small style="color: #999; font-size: 0.8em">用于购买固定资产的支出</small>
        </div>
        <div class="form-group">
          <label>营运资金变化/收入 (%)</label>
          <input v-model.number="form.wc_change_ratio" type="number" step="0.01" min="0" max="1" placeholder="0.02" />
          <small style="color: #999; font-size: 0.8em">应收账款、存货等变化</small>
        </div>
        <div class="form-group">
          <label>折旧摊销/收入 (%)</label>
          <input v-model.number="form.depreciation_ratio" type="number" step="0.01" min="0" max="1" placeholder="0.03" />
          <small style="color: #999; font-size: 0.8em">非现金支出，加回计算</small>
        </div>
        <div class="form-group">
          <label>税率 (%)</label>
          <input v-model.number="form.tax_rate" type="number" step="0.01" min="0" max="1" placeholder="0.25" />
          <small style="color: #999; font-size: 0.8em">企业所得税率</small>
        </div>
      </div>
    </div>

    <div class="form-card">
      <div class="section-title">预测参数 - WACC (加权平均资本成本)</div>
      <div class="form-grid">
        <div class="form-group">
          <label>贝塔系数 (β)</label>
          <input v-model.number="form.beta" type="number" step="0.1" placeholder="1.2" />
          <small style="color: #999; font-size: 0.8em">系统性风险指标</small>
        </div>
        <div class="form-group">
          <label>无风险利率</label>
          <input v-model.number="form.risk_free_rate" type="number" step="0.01" placeholder="0.03" />
          <small style="color: #999; font-size: 0.8em">通常使用国债收益率</small>
        </div>
        <div class="form-group">
          <label>市场风险溢价</label>
          <input v-model.number="form.market_risk_premium" type="number" step="0.01" placeholder="0.07" />
          <small style="color: #999; font-size: 0.8em">股票市场相对无风险的超额收益</small>
        </div>
        <div class="form-group">
          <label>债务成本</label>
          <input v-model.number="form.cost_of_debt" type="number" step="0.01" placeholder="0.05" />
          <small style="color: #999; font-size: 0.8em">借款利率</small>
        </div>
        <div class="form-group">
          <label>目标债务比率 (%)</label>
          <input v-model.number="form.target_debt_ratio" type="number" step="0.01" min="0" max="1" placeholder="0.3" />
          <small style="color: #999; font-size: 0.8em">目标资本结构中债务占比</small>
        </div>
      </div>
    </div>

    <div class="form-card">
      <div class="section-title">预测参数 - 终值</div>
      <div class="form-grid">
        <div class="form-group">
          <label>永续增长率 (%)</label>
          <input v-model.number="form.terminal_growth_rate" type="number" step="0.005" placeholder="0.025" />
          <small style="color: #999; font-size: 0.8em">预测期后的长期增长率</small>
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
  cost_of_debt: 0.05,
  target_debt_ratio: 0.3,
  tax_rate: 0.25,
  capex_ratio: 0.05,
  wc_change_ratio: 0.02,
  depreciation_ratio: 0.03,
  terminal_growth_rate: 0.025
} as any)

const comparables = ref<any[]>([])
const loading = ref(false)
const error = ref('')

// 多产品估值模式
const valuationMode = ref<'single' | 'multi'>('single') // 单产品/多产品模式

// 多产品数据
const products = ref<any[]>([
  {
    name: '产品A',
    description: '',
    current_revenue: 50000,
    revenue_weight: 0.6,
    growth_rate_years: [25, 20, 15, 10, 8],
    terminal_growth_rate: 3,
    gross_margin: 65,
    operating_margin: 25,
    capex_ratio: 5,
    wc_change_ratio: 2,
    depreciation_ratio: 3,
    beta: null
  }
])

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
  console.log('当前模式:', valuationMode.value)
  console.log('当前表单数据:', JSON.parse(JSON.stringify(form.value)))

  error.value = ''
  loading.value = true

  console.log('loading.value已设为true,按钮应该显示"计算中..."')

  // 重要：检查可比公司数据
  console.log('========== 开始估值检查 ==========')
  console.log('估值模式:', valuationMode.value)
  console.log('可比公司数量:', comparables.value.length)
  console.log('可比公司数据:', comparables.value)

  try {
    // 多产品估值模式
    if (valuationMode.value === 'multi') {
      console.log('========== 使用多产品估值模式 ==========')

      // 验证产品数据
      if (!validateProducts()) {
        loading.value = false
        return
      }

      // 准备多产品估值请求数据
      const multiProductRequest = {
        company_name: form.value.name,
        industry: form.value.industry,
        products: products.value.map(p => ({
          name: p.name,
          description: p.description || undefined,
          current_revenue: p.current_revenue,
          revenue_weight: p.revenue_weight,
          growth_rate_years: p.growth_rate_years.map((g: number) => g / 100),
          terminal_growth_rate: p.terminal_growth_rate / 100,
          gross_margin: p.gross_margin / 100,
          operating_margin: p.operating_margin / 100,
          capex_ratio: p.capex_ratio / 100,
          wc_change_ratio: p.wc_change_ratio / 100,
          depreciation_ratio: p.depreciation_ratio / 100,
          beta: p.beta
        })),
        tax_rate: 0.25, // 可以从表单添加
        risk_free_rate: form.value.risk_free_rate,
        market_risk_premium: form.value.market_risk_premium,
        cost_of_debt: 0.05,
        target_debt_ratio: 0.3,
        total_debt: form.value.total_debt,
        cash_and_equivalents: form.value.cash_and_equivalents,
        projection_years: 5,
        terminal_method: 'perpetuity',
        company_beta: form.value.beta
      }

      console.log('多产品估值请求:', multiProductRequest)

      // 调用多产品估值API
      const multiProductResult = await valuationAPI.multiProductDCF(multiProductRequest)

      console.log('多产品估值结果:', multiProductResult.data)

      if (!multiProductResult.data?.success) {
        throw new Error('多产品估值失败')
      }

      // 存储结果到sessionStorage（包含相对估值）
      const resultsToStore: any = {
        multiProduct: multiProductResult.data.result,
        valuationMode: 'multi',
        company: form.value,
        products: products.value,
        // 如果有可比公司，也获取相对估值结果
        ...(comparables.value.length > 0 ? {
          comparables: comparables.value,
          hasComparables: true
        } : {})
      }

      // 如果有可比公司，并行获取相对估值
      // 捕获当前comparables的快照，避免reactive引用问题
      const comparablesSnapshot = [...comparables.value]
      console.log('========== 多产品相对估值检查 ==========')
      console.log('comparables.value长度:', comparables.value.length)
      console.log('comparablesSnapshot长度:', comparablesSnapshot.length)
      console.log('comparables.value内容:', comparables.value)
      console.log('comparablesSnapshot内容:', comparablesSnapshot)

      if (comparablesSnapshot.length > 0) {
        console.log('✅ 检测到可比公司，数量:', comparablesSnapshot.length)
        console.log('开始获取相对估值结果...')
        try {
          // 准备公司数据用于相对估值
          // 准备公司数据用于相对估值
          // 计算多产品的汇总财务数据
          const totalRevenue = products.value.reduce((sum, p) => sum + p.current_revenue, 0)
          const totalNetIncome = products.value.reduce((sum, p) =>
            sum + p.current_revenue * p.operating_margin / 100, 0) * (1 - 0.25) // 估算净利润（税率25%）
          const totalEBITDA = products.value.reduce((sum, p) =>
            sum + p.current_revenue * p.operating_margin / 100 * 1.2, 0) // 粗略估算EBITDA

          // 构造完整的公司对象，符合后端CompanyInput模型要求
          const companyForRelative = {
            ...form.value,  // 包含所有必需字段（name, industry, stage, beta等）
            revenue: totalRevenue,
            net_income: totalNetIncome,
            ebitda: totalEBITDA,
            net_assets: form.value.net_assets,
            total_debt: form.value.total_debt,
            cash_and_equivalents: form.value.cash_and_equivalents,
            growth_rate: totalRevenue / form.value.revenue - 1, // 基于总收入增长计算
            operating_margin: totalNetIncome / totalRevenue, // 综合营业利润率
          }

          console.log('多产品模式相对估值请求参数:', companyForRelative)
          console.log('可比公司数据:', comparablesSnapshot)
          console.log('即将调用valuationAPI.relative，参数:', {
            company: companyForRelative,
            comparables: comparablesSnapshot,
            methods: undefined
          })

          const relativeResult = await valuationAPI.relative(companyForRelative, comparablesSnapshot)
          console.log('相对估值API完整响应:', relativeResult)
          console.log('相对估值API响应.data:', relativeResult.data)
          console.log('相对估值API响应.data.success:', relativeResult.data?.success)
          console.log('相对估值API响应.status:', relativeResult.status)
          console.log('相对估值API响应.statusText:', relativeResult.statusText)

          if (relativeResult.data?.success) {
            resultsToStore.relative = relativeResult.data  // 存储整个data对象，与单产品模式保持一致
            console.log('✅ 相对估值数据已添加到结果中, resultsToStore.relative:', resultsToStore.relative)
          } else {
            console.warn('⚠️ 相对估值API返回success=false或无success字段')
            console.warn('响应数据:', relativeResult.data)
            // 存储错误信息用于调试
            resultsToStore.relativeError = {
              message: 'API返回success=false',
              response: relativeResult.data
            }
          }
        } catch (relErr) {
          console.error('❌ 相对估值获取异常:', relErr)
          console.error('错误详情:', (relErr as any).message)
          console.error('错误堆栈:', (relErr as any).stack)
          // 存储异常信息用于调试
          resultsToStore.relativeError = {
            message: (relErr as any).message,
            stack: (relErr as any).stack
          }
          // 不抛出错误，允许多产品估值继续进行
        }
      } else {
        console.log('ℹ️ 未添加可比公司，跳过相对估值')
        resultsToStore.noComparablesReason = 'comparablesSnapshot.length为0'
      }

      console.log('准备存储到sessionStorage的数据:', resultsToStore)

      const jsonStr = JSON.stringify(resultsToStore)
      sessionStorage.setItem('valuationResults', jsonStr)

      await nextTick()
      router.push('/valuation/result')
      console.log('✅ 多产品估值完成，跳转到结果页')

      return
    }

    // 单产品估值模式（原有逻辑）
    const company = {
      ...form.value,
      growth_rate: form.value.growth_rate / 100,
      operating_margin: form.value.operating_margin / 100,
      target_debt_ratio: form.value.target_debt_ratio || 0.3,
      tax_rate: form.value.tax_rate || 0.25,
      capex_ratio: form.value.capex_ratio || 0.05,
      wc_change_ratio: form.value.wc_change_ratio || 0.02,
      depreciation_ratio: form.value.depreciation_ratio || 0.03,
      cost_of_debt: form.value.cost_of_debt || 0.05
    }

    console.log('公司数据:', company)
    console.log('可比公司数量:', comparables.value.length)

    // 并行执行多个估值请求
    // 情景分析使用默认参数（基准、乐观、悲观）
    const requests = [
      valuationAPI.dcf(company),
      scenarioAPI.analyze(company),  // 使用默认情景
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
    // 需要存储转换后的 company 对象（growth_rate 和 operating_margin 已除以100）
    const companyForStorage = {
      ...form.value,
      growth_rate: form.value.growth_rate / 100,
      operating_margin: form.value.operating_margin / 100
    }
    const resultsToStore = {
      relative: relativeResult?.data,
      dcf: dcfResult?.data,  // dcfResult.data = {success: true, result: {...}}
      scenario: scenarioResult?.data,
      stress: stressResult?.data,
      sensitivity: sensitivityResult?.data,
      company: companyForStorage,
      comparables: comparables.value,
      valuationMode: 'single'
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
    cost_of_debt: 0.05,
    target_debt_ratio: 0.3,
    tax_rate: 0.25,
    capex_ratio: 0.05,
    wc_change_ratio: 0.02,
    depreciation_ratio: 0.03,
    terminal_growth_rate: 0.025
  }
  comparables.value = []
}

// ===== 多产品管理函数 =====

// 添加产品
const addProduct = () => {
  if (products.value.length >= 10) {
    error.value = '最多只能添加10个产品'
    return
  }

  const newProduct = {
    name: `产品${String.fromCharCode(65 + products.value.length)}`,
    description: '',
    current_revenue: 0,
    revenue_weight: 0,
    growth_rate_years: [15, 15, 15, 15, 15],
    terminal_growth_rate: 2.5,
    gross_margin: 50,
    operating_margin: 20,
    capex_ratio: 5,
    wc_change_ratio: 2,
    depreciation_ratio: 3,
    beta: null
  }

  products.value.push(newProduct)
  updateProductWeights()
}

// 删除产品
const removeProduct = (index: number) => {
  if (products.value.length <= 1) {
    error.value = '至少需要保留1个产品'
    return
  }
  products.value.splice(index, 1)
  updateProductWeights()
}

// 更新产品权重（自动均分）
// 自动计算收入占比（根据当前收入）
const autoCalculateWeights = () => {
  const totalRevenue = products.value.reduce((sum, p) => sum + (p.current_revenue || 0), 0)
  if (totalRevenue > 0) {
    products.value.forEach(p => {
      p.revenue_weight = Math.round((p.current_revenue / totalRevenue) * 1000) / 1000
    })
  }
}

// 手动修改权重后标记为手动模式
const manualWeightChange = () => {
  // 权重手动修改后不再自动计算，由用户自行调整
}

const updateProductWeights = () => {
  const count = products.value.length
  const weight = 1 / count
  products.value.forEach(p => {
    p.revenue_weight = Math.round(weight * 100) / 100
  })
}

// 验证产品数据
const validateProducts = () => {
  // 检查权重总和
  const totalWeight = products.value.reduce((sum, p) => sum + (p.revenue_weight || 0), 0)
  if (Math.abs(totalWeight - 1) > 0.01) {
    error.value = `产品权重总和应为100%，当前为${(totalWeight * 100).toFixed(1)}%`
    return false
  }

  // 检查每个产品的必填字段
  for (let i = 0; i < products.value.length; i++) {
    const p = products.value[i]
    if (!p.name || p.name.trim() === '') {
      error.value = `产品${i + 1}的名称不能为空`
      return false
    }
    if (!p.current_revenue || p.current_revenue <= 0) {
      error.value = `产品${p.name}的当前收入必须大于0`
      return false
    }
    if (!p.revenue_weight || p.revenue_weight <= 0 || p.revenue_weight > 1) {
      error.value = `产品${p.name}的权重必须在0-100%之间`
      return false
    }
  }

  return true
}

// 计算总权重
const totalWeight = computed(() => {
  return products.value.reduce((sum, p) => sum + (p.revenue_weight || 0), 0)
})
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

.btn-auto-calc {
  background: #10b981;
  color: white;
  border: none;
  padding: 6px 12px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 0.85em;
  margin-left: auto;
}

.btn-auto-calc:hover {
  background: #059669;
}

.form-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
  gap: 20px;
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

/* 公司基本信息对齐优化 */
.company-info-grid {
  grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
  gap: 20px;
}

.company-info-grid .form-group {
  display: flex;
  flex-direction: column;
}

/* 所属行业占据整行 */
.company-info-grid .form-group-full-width {
  grid-column: 1 / -1;  /* 跨越所有列 */
}

/* 行业选择特殊处理 */
.company-info-grid .form-group .industry-cascade {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.company-info-grid .form-group .industry-l3-row {
  width: 100%;
  margin-top: 10px;
  display: flex;
}

.industry-select-wide {
  flex: 1;
  min-width: 200px;
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
  gap: 20px;
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
  grid-template-columns: repeat(2, 1fr);  /* 固定2列 */
  gap: 10px;
  margin-top: 10px;
}

@media (max-width: 768px) {
  .filter-grid {
    grid-template-columns: 1fr;  /* 小屏幕单列 */
  }
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

/* 估值模式切换 */
.mode-switch {
  display: flex;
  gap: 15px;
  margin-bottom: 15px;
}

.mode-btn {
  flex: 1;
  padding: 12px 24px;
  border: 2px solid #ddd;
  background: white;
  border-radius: 8px;
  font-size: 16px;
  font-weight: 500;
  color: #666;
  cursor: pointer;
  transition: all 0.3s;
}

.mode-btn:hover {
  border-color: #667eea;
  color: #667eea;
}

.mode-btn.active {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  border-color: #667eea;
  color: white;
}

.mode-description {
  padding: 12px 16px;
  background: #f8f9fa;
  border-radius: 6px;
  color: #666;
  font-size: 0.95em;
  line-height: 1.6;
}

/* 多产品配置 */
.weight-indicator {
  margin-left: auto;
  padding: 6px 12px;
  background: #e8f4ff;
  border-radius: 4px;
  color: #667eea;
  font-weight: 600;
}

.weight-indicator.error {
  background: #ffe8e8;
  color: #dc3545;
}

.products-header {
  display: flex;
  align-items: center;
  gap: 20px;
  margin-bottom: 20px;
  padding-bottom: 15px;
  border-bottom: 1px solid #eee;
}

.products-count {
  color: #666;
  font-size: 0.95em;
}

.products-list {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.product-card {
  background: #fafafa;
  border: 1px solid #e0e0e0;
  border-radius: 10px;
  padding: 20px;
  transition: box-shadow 0.3s;
}

.product-card:hover {
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.product-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 15px;
  padding-bottom: 10px;
  border-bottom: 1px solid #ddd;
}

.product-title {
  font-size: 1.1em;
  font-weight: 600;
  color: #333;
}

.btn-remove-product {
  padding: 6px 12px;
  background: #fff;
  border: 1px solid #dc3545;
  color: #dc3545;
  border-radius: 4px;
  cursor: pointer;
  transition: all 0.3s;
}

.btn-remove-product:hover {
  background: #dc3545;
  color: white;
}

.product-section {
  margin: 20px 0;
  padding: 15px;
  background: white;
  border-radius: 8px;
  border: 1px solid #e0e0e0;
}

.product-section-title {
  font-weight: 600;
  color: #555;
  margin-bottom: 15px;
}

.growth-years-input {
  display: grid;
  grid-template-columns: repeat(2, 1fr);  /* 固定2列 */
  gap: 20px;
}

@media (max-width: 768px) {
  .growth-years-input {
    grid-template-columns: 1fr;  /* 小屏幕单列 */
  }
}

.year-input {
  display: flex;
  flex-direction: column;
  gap: 5px;
}

.year-input label {
  font-size: 0.85em;
  color: #666;
}

.year-input input {
  width: 100%;
  padding: 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  text-align: center;
}

.year-input span {
  text-align: center;
  font-size: 0.85em;
  color: #999;
}

.btn-remove-product {
  padding: 6px 12px;
  background: #fff;
  border: 1px solid #dc3545;
  color: #dc3545;
  border-radius: 4px;
  cursor: pointer;
  transition: all 0.3s;
}

.btn-remove-product:hover {
  background: #dc3545;
  color: white;
}
</style>
