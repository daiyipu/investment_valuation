<script setup lang="ts">
import { RouterLink, RouterView } from 'vue-router'
</script>

<template>
  <div class="app">
    <nav class="navbar">
      <div class="nav-container">
        <RouterLink to="/" class="nav-brand">
          🏦 投资估值系统
        </RouterLink>
        <div class="nav-links">
          <RouterLink to="/" class="nav-link">首页</RouterLink>
          <RouterLink to="/valuation" class="nav-link">估值</RouterLink>
          <RouterLink to="/history" class="nav-link">历史记录</RouterLink>
          <RouterLink to="/scenario" class="nav-link">情景分析</RouterLink>
          <RouterLink to="/sensitivity" class="nav-link">敏感性</RouterLink>
          <RouterLink to="/stress-test" class="nav-link">压力测试</RouterLink>
          <RouterLink to="/report" class="nav-link">报告</RouterLink>
        </div>
      </div>
    </nav>

    <main class="main-content">
      <RouterView />
    </main>

    <footer class="footer">
      <p>&copy; 2024 投资估值系统 v1.0 | 股权投资基金专业工具</p>
    </footer>
  </div>
</template>

<style scoped>
.app {
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  background: #f5f7fa;
}

.navbar {
  background: white;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  position: sticky;
  top: 0;
  z-index: 100;
}

.nav-container {
  max-width: 1400px;
  margin: 0 auto;
  padding: 0 20px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  height: 60px;
}

.nav-brand {
  font-size: 1.2em;
  font-weight: bold;
  color: #667eea;
  text-decoration: none;
  white-space: nowrap;
}

.nav-links {
  display: flex;
  gap: 5px;
}

.nav-link {
  padding: 8px 16px;
  color: #555;
  text-decoration: none;
  border-radius: 6px;
  transition: all 0.3s;
  font-size: 0.9em;
  white-space: nowrap;
}

.nav-link:hover {
  background: #f0f0f0;
  color: #667eea;
}

.nav-link.router-link-active {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
}

.main-content {
  flex: 1;
  padding: 20px 0;
}

.footer {
  background: white;
  padding: 20px;
  text-align: center;
  color: #666;
  font-size: 0.9em;
  border-top: 1px solid #e0e0e0;
}

@media (max-width: 768px) {
  .nav-container {
    flex-direction: column;
    height: auto;
    padding: 10px 20px;
  }

  .nav-brand {
    margin-bottom: 10px;
  }

  .nav-links {
    flex-wrap: wrap;
    justify-content: center;
  }
}
</style>

<style>
/* Global styles */
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: 'Microsoft YaHei', 'PingFang SC', 'Helvetica Neue', Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

#app {
  min-height: 100vh;
}
</style>
