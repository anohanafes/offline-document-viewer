import 'core-js/stable'
import Vue from 'vue'
import App from './App.vue'
import router from './router'

// 导入全局样式
import './assets/styles/main.css'
// 导入原项目的Excel样式
import './assets/styles/excel-styles.css'

Vue.config.productionTip = false

// 全局错误处理
Vue.config.errorHandler = (err, vm, info) => {
  console.error('Vue错误:', err)
  console.error('错误信息:', info)
}

// 全局组件注册（如果需要的话）
// Vue.component('GlobalComponent', GlobalComponent)

new Vue({
  router,
  render: h => h(App)
}).$mount('#app')
