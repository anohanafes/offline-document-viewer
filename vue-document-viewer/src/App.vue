<template>
  <div id="app">
    <nav class="app-nav" v-if="showNavigation">
      <div class="nav-container">
        <!-- <router-link to="/local" class="nav-brand">
          📄 Vue文档预览器
        </router-link> -->
        <div>
          📄 Vue文档预览器
        </div>
        <div class="nav-links">
          <router-link to="/local" class="nav-link">本地文件</router-link>
          <router-link to="/url" class="nav-link">URL预览</router-link>
          <router-link to="/direct" class="nav-link">直接预览</router-link>
        </div>
      </div>
    </nav>

    <main class="app-main" :class="{ 'no-nav': !showNavigation }">
      <router-view />
    </main>

    <!-- 全局加载遮罩 -->
    <div v-if="globalLoading" class="global-loading">
      <div class="loading-spinner"></div>
      <div class="loading-text">{{ loadingText }}</div>
    </div>
  </div>
</template>

<script>
export default {
  name: 'App',
  data() {
    return {
      globalLoading: false,
      loadingText: '加载中...'
    }
  },
  computed: {
    showNavigation() {
      // 在direct模式下隐藏导航栏
      return this.$route.name !== 'DirectViewer'
    }
  },
  methods: {
    setGlobalLoading(loading, text = '加载中...') {
      this.globalLoading = loading
      this.loadingText = text
    }
  },
  created() {
    // 提供全局加载状态管理
    this.$root.setGlobalLoading = this.setGlobalLoading
  }
}
</script>

<style>
/* 全局样式重置 */
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

html, body {
  height: 100%;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', 'Helvetica Neue', Helvetica, Arial, sans-serif;
  background-color: #f5f5f5;
}

#app {
  height: 100vh;
  display: flex;
  flex-direction: column;
}

.app-nav {
  background: #fff;
  border-bottom: 1px solid #e0e0e0;
  z-index: 100;
}

.app-nav .nav-container {
  max-width: 1200px;
  margin: 0 auto;
  padding: 0 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  height: 60px;
}

.app-nav .nav-brand {
  font-size: 20px;
  font-weight: 600;
  color: #2196F3;
  text-decoration: none;
}

.app-nav .nav-brand:hover {
  color: #1976D2;
}

.app-nav .nav-links {
  display: flex;
  gap: 20px;
}

.app-nav .nav-link {
  color: #666;
  text-decoration: none;
  padding: 8px 16px;
  border-radius: 4px;
  transition: all 0.3s;
}

.app-nav .nav-link:hover {
  background: #f0f8ff;
  color: #2196F3;
}

.app-nav .nav-link.router-link-active {
  background: #2196F3;
  color: white;
}

.app-main {
  flex: 1;
  overflow: hidden;
}

.app-main.no-nav {
  height: 100vh;
}

.global-loading {
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  z-index: 9999;
}

.global-loading .loading-spinner {
  width: 40px;
  height: 40px;
  border: 4px solid rgba(255, 255, 255, 0.3);
  border-top: 4px solid #2196F3;
  border-radius: 50%;
  animation: spin 1s linear infinite;
}

.global-loading .loading-text {
  color: white;
  margin-top: 16px;
  font-size: 16px;
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

/* 响应式设计 */
@media (max-width: 768px) {
  .app-nav .nav-container {
    padding: 0 16px;
    height: 50px;
  }
  
  .app-nav .nav-brand {
    font-size: 18px;
  }
  
  .app-nav .nav-links {
    gap: 12px;
  }
  
  .app-nav .nav-link {
    padding: 6px 12px;
    font-size: 14px;
  }
}
</style>
