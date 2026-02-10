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
          <label>所属行业</label>
          <select v-model="form.industry">
            <option value="">请选择行业...</option>
            <optgroup label="Tushare行业分类">
              <option value="软件服务">软件服务</option>
              <option value="互联网">互联网</option>
              <option value="IT设备">IT设备</option>
              <option value="通信设备">通信设备</option>
              <option value="半导体">半导体</option>
              <option value="元器件">元器件</option>
              <option value="电器仪表">电器仪表</option>
              <option value="电气设备">电气设备</option>
            </optgroup>
            <optgroup label="医药生物">
              <option value="生物制药">生物制药</option>
              <option value="化学制药">化学制药</option>
              <option value="中成药">中成药</option>
              <option value="医药商业">医药商业</option>
              <option value="医疗保健">医疗保健</option>
            </optgroup>
            <optgroup label="消费">
              <option value="食品">食品</option>
              <option value="白酒">白酒</option>
              <option value="红黄酒">红黄酒</option>
              <option value="软饮料">软饮料</option>
              <option value="乳制品">乳制品</option>
              <option value="纺织">纺织</option>
              <option value="服饰">服饰</option>
              <option value="家居用品">家居用品</option>
              <option value="家用电器">家用电器</option>
              <option value="文教休闲">文教休闲</option>
              <option value="影视音像">影视音像</option>
              <option value="广告包装">广告包装</option>
              <option value="出版业">出版业</option>
              <option value="商贸代理">商贸代理</option>
              <option value="百货">百货</option>
              <option value="超市连锁">超市连锁</option>
              <option value="酒店餐饮">酒店餐饮</option>
              <option value="旅游景点">旅游景点</option>
              <option value="旅游服务">旅游服务</option>
            </optgroup>
            <optgroup label="金融">
              <option value="银行">银行</option>
              <option value="证券">证券</option>
              <option value="保险">保险</option>
              <option value="多元金融">多元金融</option>
            </optgroup>
            <optgroup label="房地产">
              <option value="全国地产">全国地产</option>
              <option value="区域地产">区域地产</option>
              <option value="房产服务">房产服务</option>
              <option value="园区开发">园区开发</option>
            </optgroup>
            <optgroup label="建筑建材">
              <option value="建筑工程">建筑工程</option>
              <option value="建筑装饰">建筑装饰</option>
              <option value="其他建材">其他建材</option>
              <option value="水泥">水泥</option>
              <option value="玻璃">玻璃</option>
            </optgroup>
            <optgroup label="汽车">
              <option value="汽车整车">汽车整车</option>
              <option value="汽车配件">汽车配件</option>
              <option value="汽车服务">汽车服务</option>
              <option value="摩托车">摩托车</option>
              <option value="其他交运设备">其他交运设备</option>
            </optgroup>
            <optgroup label="机械制造">
              <option value="工程机械">工程机械</option>
              <option value="专用机械">专用机械</option>
              <option value="通用机械">通用机械</option>
              <option value="机械基件">机械基件</option>
              <option value="机床制造">机床制造</option>
              <option value="轻工机械">轻工机械</option>
              <option value="纺织机械">纺织机械</option>
              <option value="运输设备">运输设备</option>
            </optgroup>
            <optgroup label="化工">
              <option value="化工原料">化工原料</option>
              <option value="化工机械">化工机械</option>
              <option value="化纤">化纤</option>
              <option value="塑料">塑料</option>
              <option value="橡胶">橡胶</option>
              <option value="日用化工">日用化工</option>
              <option value="农药化肥">农药化肥</option>
              <option value="染料涂料">染料涂料</option>
            </optgroup>
            <optgroup label="能源电力">
              <option value="电力设备">电力设备</option>
              <option value="新型电力">新型电力</option>
              <option value="火力发电">火力发电</option>
              <option value="水力发电">水力发电</option>
              <option value="水务">水务</option>
              <option value="供气供热">供气供热</option>
              <option value="石油开采">石油开采</option>
              <option value="石油加工">石油加工</option>
              <option value="石油贸易">石油贸易</option>
              <option value="煤炭开采">煤炭开采</option>
              <option value="焦炭加工">焦炭加工</option>
              <option value="燃气">燃气</option>
            </optgroup>
            <optgroup label="金属采矿">
              <option value="钢铁">钢铁</option>
              <option value="特钢">特种钢</option>
              <option value="普钢">普钢</option>
              <option value="有色金属">有色金属</option>
              <option value="铜">铜</option>
              <option value="铝">铝</option>
              <option value="铅锌">铅锌</option>
              <option value="黄金">黄金</option>
              <option value="小金属">小金属</option>
              <option value="矿物制品">矿物制品</option>
            </optgroup>
            <optgroup label="交通运输">
              <option value="铁路">铁路</option>
              <option value="公路">公路</option>
              <option value="水路">水路</option>
              <option value="港口">港口</option>
              <option value="机场">机场</option>
              <option value="航空">航空</option>
              <option value="航运">航运</option>
              <option value="空运">空运</option>
            </optgroup>
            <optgroup label="其他">
              <option value="综合类">综合类</option>
              <option value="其他商业">其他商业</option>
            </optgroup>
          </select>
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
      <div class="section-title">财务数据（单位：万元）</div>
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
        可比公司数据（可选，用于相对估值）
      </div>

      <!-- 导入选项 -->
      <div class="import-options">
        <button @click="showImportModal = true" class="btn-import" type="button">
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
        <p>暂无可比公司数据，将仅使用DCF估值</p>
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
            <label>选择行业</label>
            <select v-model="importIndustry" @change="onIndustryChange">
              <option value="">请选择行业...</option>
              <optgroup label="热门行业">
                <option value="软件服务">软件服务</option>
                <option value="互联网">互联网</option>
                <option value="IT设备">IT设备</option>
                <option value="通信设备">通信设备</option>
                <option value="半导体">半导体</option>
                <option value="元器件">元器件</option>
                <option value="电气设备">电气设备</option>
                <option value="生物制药">生物制药</option>
                <option value="化学制药">化学制药</option>
                <option value="中成药">中成药</option>
                <option value="医药商业">医药商业</option>
                <option value="医疗保健">医疗保健</option>
                <option value="汽车整车">汽车整车</option>
                <option value="汽车配件">汽车配件</option>
                <option value="银行">银行</option>
                <option value="证券">证券</option>
                <option value="保险">保险</option>
              </optgroup>
              <optgroup label="TMT科技">
                <option value="IT设备">IT设备</option>
                <option value="互联网">互联网</option>
                <option value="半导体">半导体</option>
                <option value="元器件">元器件</option>
                <option value="通信设备">通信设备</option>
                <option value="软件服务">软件服务</option>
                <option value="电信运营">电信运营</option>
              </optgroup>
              <optgroup label="医药生物">
                <option value="生物制药">生物制药</option>
                <option value="化学制药">化学制药</option>
                <option value="中成药">中成药</option>
                <option value="医药商业">医药商业</option>
                <option value="医疗保健">医疗保健</option>
                <option value="医疗器械">医疗器械</option>
                <option value="医疗服务">医疗服务</option>
              </optgroup>
              <optgroup label="消费零售">
                <option value="百货">百货</option>
                <option value="超市连锁">超市连锁</option>
                <option value="商贸代理">商贸代理</option>
                <option value="纺织">纺织</option>
                <option value="服饰">服饰</option>
                <option value="家居">家居</option>
                <option value="旅游服务">旅游服务</option>
                <option value="旅游景点">旅游景点</option>
                <option value="酒店餐饮">酒店餐饮</option>
                <option value="农业综合">农业综合</option>
                <option value="农产品加工">农产品加工</option>
                <option value="食品加工">食品加工</option>
              </optgroup>
              <optgroup label="汽车制造">
                <option value="汽车整车">汽车整车</option>
                <option value="汽车配件">汽车配件</option>
                <option value="摩托车">摩托车</option>
                <option value="其他交运设备">其他交运设备</option>
              </optgroup>
              <optgroup label="机械设备">
                <option value="机械基件">机械基件</option>
                <option value="机床制造">机床制造</option>
                <option value="电气设备">电气设备</option>
                <option value="电器仪表">电器仪表</option>
                <option value="运输设备">运输设备</option>
                <option value="重工机械">重工机械</option>
                <option value="农用机械">农用机械</option>
              </optgroup>
              <optgroup label="房地产建材">
                <option value="全国地产">全国地产</option>
                <option value="区域地产">区域地产</option>
                <option value="房产服务">房产服务</option>
                <option value="园区开发">园区开发</option>
                <option value="房屋建设">房屋建设</option>
                <option value="建筑装饰">建筑装饰</option>
                <option value="建筑材料">建筑材料</option>
                <option value="水泥建材">水泥建材</option>
                <option value="其他建材">其他建材</option>
              </optgroup>
              <optgroup label="化工能源">
                <option value="基础化工">基础化工</option>
                <option value="化工原料">化工原料</option>
                <option value="化纤">化纤</option>
                <option value="农药化肥">农药化肥</option>
                <option value="石油">石油</option>
                <option value="煤炭">煤炭</option>
                <option value="电力">电力</option>
                <option value="供气供热">供气供热</option>
                <option value="环境保护">环境保护</option>
                <option value="新材料">新材料</option>
              </optgroup>
              <optgroup label="金融">
                <option value="银行">银行</option>
                <option value="证券">证券</option>
                <option value="保险">保险</option>
                <option value="多元金融">多元金融</option>
              </optgroup>
              <optgroup label="交通运输">
                <option value="公路">公路</option>
                <option value="机场">机场</option>
                <option value="港口">港口</option>
                <option value="水运">水运</option>
                <option value="铁路">铁路</option>
                <option value="物流">物流</option>
              </optgroup>
              <optgroup label="文化传媒教育">
                <option value="文化传媒">文化传媒</option>
                <option value="出版业">出版业</option>
                <option value="影视动漫">影视动漫</option>
                <option value="营销传播">营销传播</option>
                <option value="教育">教育</option>
                <option value="体育">体育</option>
                <option value="文教休闲">文教休闲</option>
              </optgroup>
              <optgroup label="其他">
                <option value="航空航天">航空航天</option>
                <option value="船舶制造">船舶制造</option>
                <option value="造纸">造纸</option>
                <option value="包装印刷">包装印刷</option>
                <option value="建筑施工">建筑施工</option>
                <option value="其他">其他</option>
              </optgroup>
            </select>
          </div>

          <div class="form-group">
            <label>筛选条件</label>
            <div class="filter-grid">
              <div>
                <label>最小市值（亿元）</label>
                <input v-model.number="importMinMarketCap" type="number" placeholder="不限制">
              </div>
              <div>
                <label>最大市值（亿元）</label>
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
      <button class="btn btn-primary" @click="startValuation" :disabled="loading">
        {{ loading ? '计算中...' : '🚀 开始估值' }}
      </button>
      <button class="btn btn-secondary" @click="resetForm">🔄 重置</button>
    </div>

    <div v-if="error" class="error">{{ error }}</div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'
import { useRouter } from 'vue-router'
import { valuationAPI, scenarioAPI, stressTestAPI, sensitivityAPI, dataAPI } from '../services/api'
import axios from 'axios'

const router = useRouter()

const form = ref({
  name: '云数科技有限公司',
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
})

const comparables = ref<any[]>([])
const loading = ref(false)
const error = ref('')

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

const onIndustryChange = () => {
  // 当选择行业时，如果目标公司没有行业，自动填充
  if (!form.value.industry && importIndustry.value) {
    form.value.industry = importIndustry.value
  }
}

const importFromTushare = async () => {
  if (!importIndustry.value) {
    importError.value = '请先选择行业'
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
        importError.value = '未找到符合条件的公司，请尝试调整筛选条件或选择其他行业'
      }
    }
  } catch (err: any) {
    console.error('导入失败:', err)
    importError.value = '导入失败: ' + (err.response?.data?.detail || err.message)
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
  console.error('详细错误:', err)

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
  error.value = ''
  loading.value = true

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

    // 如果有可比公司，添加相对估值
    if (comparables.value.length > 0) {
      console.log('可比公司数据:', comparables.value)
      requests.unshift(valuationAPI.relative(company, comparables.value))
    }

    const results = await Promise.all(requests)

    let dcfResult, scenarioResult, stressResult, sensitivityResult, relativeResult

    if (comparables.value.length > 0) {
      [relativeResult, dcfResult, scenarioResult, stressResult, sensitivityResult] = results
    } else {
      [dcfResult, scenarioResult, stressResult, sensitivityResult] = results
    }

    console.log('估值结果:', {
      relative: relativeResult?.data,
      dcf: dcfResult.data,
      scenario: scenarioResult.data
    })

    // 存储结果到sessionStorage用于结果页展示
    sessionStorage.setItem('valuationResults', JSON.stringify({
      relative: relativeResult?.data,
      dcf: dcfResult.data,
      scenario: scenarioResult.data,
      stress: stressResult.data,
      sensitivity: sensitivityResult.data,
      company: form.value,
      comparables: comparables.value
    }))

    router.push('/valuation/result')
  } catch (err: any) {
    error.value = '估值计算失败: ' + getErrorMessage(err)
  } finally {
    loading.value = false
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
</style>
