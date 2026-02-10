import { createRouter, createWebHistory } from 'vue-router'

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      name: 'dashboard',
      component: () => import('../views/Dashboard.vue')
    },
    {
      path: '/valuation',
      name: 'valuation',
      component: () => import('../views/ValuationInput.vue')
    },
    {
      path: '/valuation/result',
      name: 'valuation-result',
      component: () => import('../views/ValuationResult.vue')
    },
    {
      path: '/scenario',
      name: 'scenario',
      component: () => import('../views/ScenarioAnalysis.vue')
    },
    {
      path: '/stress-test',
      name: 'stress-test',
      component: () => import('../views/StressTest.vue')
    },
    {
      path: '/sensitivity',
      name: 'sensitivity',
      component: () => import('../views/SensitivityAnalysis.vue')
    },
    {
      path: '/report',
      name: 'report',
      component: () => import('../views/Report.vue')
    }
  ]
})

export default router
