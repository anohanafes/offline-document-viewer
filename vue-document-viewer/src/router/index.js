import Vue from 'vue'
import VueRouter from 'vue-router'

// 导入页面组件
import LocalFileViewer from '@/views/LocalFileViewer.vue'
import UrlViewer from '@/views/UrlViewer.vue'
import DirectViewer from '@/views/DirectViewer.vue'

Vue.use(VueRouter)

const routes = [
  {
    path: '/local',
    name: 'LocalFileViewer',
    component: LocalFileViewer,
    meta: {
      title: '本地文件预览',
      description: '选择本地文件进行预览'
    }
  },
  {
    path: '/url',
    name: 'UrlViewer',
    component: UrlViewer,
    meta: {
      title: 'URL文档预览',
      description: '通过URL链接预览在线文档'
    }
  },
  {
    path: '/direct',
    name: 'DirectViewer',
    component: DirectViewer,
    meta: {
      title: '直接预览',
      description: '嵌入式文档预览模式'
    }
  },
  {
    // 404 重定向
    path: '*',
    redirect: '/'
  }
]

const router = new VueRouter({
  mode: 'hash', // 使用hash模式，方便部署
  base: process.env.BASE_URL,
  routes
})

// 路由守卫
router.beforeEach((to, from, next) => {
  // 设置页面标题
  if (to.meta.title) {
    document.title = `${to.meta.title} - Vue文档预览器`
  }
  next()
})

export default router
