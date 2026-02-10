<template>
  <div class="dashboard">
    <div class="header">
      <h1>🏦 投资估值系统</h1>
      <p>专业股权投资基金估值工具 v1.0</p>
    </div>

    <div class="cards-grid">
      <div class="card" @click="$router.push('/valuation')">
        <div class="card-icon">💰</div>
        <div class="card-title">相对估值</div>
        <div class="card-desc">P/E、P/S、P/B、EV/EBITDA等方法</div>
      </div>

      <div class="card" @click="$router.push('/valuation')">
        <div class="card-icon">📊</div>
        <div class="card-title">绝对估值</div>
        <div class="card-desc">DCF现金流折现模型</div>
      </div>

      <div class="card" @click="$router.push('/scenario')">
        <div class="card-icon">📈</div>
        <div class="card-title">情景分析</div>
        <div class="card-desc">基准/乐观/悲观情景对比</div>
      </div>

      <div class="card" @click="$router.push('/stress-test')">
        <div class="card-icon">⚠️</div>
        <div class="card-title">压力测试</div>
        <div class="card-desc">收入冲击、毛利率压缩、蒙特卡洛</div>
      </div>

      <div class="card" @click="$router.push('/sensitivity')">
        <div class="card-icon">📉</div>
        <div class="card-title">敏感性分析</div>
        <div class="card-desc">单因素/双因素敏感性、龙卷风图</div>
      </div>

      <div class="card" @click="$router.push('/report')">
        <div class="card-icon">📄</div>
        <div class="card-title">综合报告</div>
        <div class="card-desc">完整估值分析报告</div>
      </div>
    </div>

    <div class="features">
      <h2>系统功能</h2>
      <ul>
        <li>✅ 多种估值方法交叉验证</li>
        <li>✅ 实时获取A股可比公司数据（Tushare）</li>
        <li>✅ 情景分析与压力测试</li>
        <li>✅ 可视化图表展示</li>
        <li>✅ 一键生成综合报告</li>
      </ul>
    </div>
  </div>
</template>

<script setup lang="ts">
</script>

<style scoped>
.dashboard {
  padding: 20px;
  max-width: 1400px;
  margin: 0 auto;
}

.header {
  text-align: center;
  margin-bottom: 40px;
  padding: 40px 20px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  border-radius: 12px;
  color: white;
}

.header h1 {
  font-size: 2.5em;
  margin-bottom: 10px;
}

.header p {
  font-size: 1.1em;
  opacity: 0.9;
}

.cards-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 20px;
  margin-bottom: 40px;
}

.card {
  background: white;
  padding: 30px;
  border-radius: 12px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
  cursor: pointer;
  transition: all 0.3s;
  text-align: center;
}

.card:hover {
  transform: translateY(-5px);
  box-shadow: 0 8px 24px rgba(102, 126, 234, 0.3);
}

.card-icon {
  font-size: 3em;
  margin-bottom: 15px;
}

.card-title {
  font-size: 1.3em;
  font-weight: bold;
  color: #333;
  margin-bottom: 10px;
}

.card-desc {
  color: #666;
  font-size: 0.9em;
  line-height: 1.5;
}

.features {
  background: white;
  padding: 30px;
  border-radius: 12px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.features h2 {
  color: #333;
  margin-bottom: 20px;
  font-size: 1.5em;
}

.features ul {
  list-style: none;
  padding: 0;
}

.features li {
  padding: 10px 0;
  font-size: 1.1em;
  color: #555;
  border-bottom: 1px solid #f0f0f0;
}

.features li:last-child {
  border-bottom: none;
}
</style>
