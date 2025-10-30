<template>
  <div class="direct-viewer">
    <!-- 全屏文档预览 -->
    <div v-if="documentLoaded" class="document-section">
      <DocumentViewer
        :file="documentFile"
        :file-name="documentFileName"
        :hide-header="true"
        :hide-controls="false"
        @document-loaded="handleDocumentLoaded"
        @document-error="handleDocumentError"
        ref="documentViewer"
      />
    </div>

    <!-- 初始加载界面 -->
    <div v-else-if="!error" class="loading-section">
      <div class="loading-content">
        <div class="loading-icon">📄</div>
        <h2>直接预览模式</h2>
        <p v-if="!documentUrl">请在URL中提供文档地址</p>
        <p v-else>正在加载文档...</p>
        
        <div v-if="loading" class="loading-progress">
          <div class="loading-spinner"></div>
          <div class="loading-text">{{ loadingText }}</div>
        </div>
      </div>
    </div>

    <!-- 错误界面 -->
    <div v-if="error" class="error-section">
      <div class="error-content">
        <div class="error-icon">❌</div>
        <h2>加载失败</h2>
        <p class="error-message">{{ error }}</p>
        
        <div class="error-actions">
          <button @click="retry" class="btn btn-primary">重试</button>
          <button @click="goBack" class="btn btn-secondary">返回</button>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import DocumentViewer from '@/components/DocumentViewer.vue'
import { DocumentUtils } from '@/utils/documentUtils'

export default {
  name: 'DirectViewer',
  components: {
    DocumentViewer
  },
  data() {
    return {
      documentLoaded: false,
      documentFile: null,
      documentFileName: '',
      
      loading: false,
      loadingText: '正在加载文档...',
      
      error: null,
      errorDetails: null
    }
  },
  computed: {
    documentUrl() {
      let url = this.$route.query.url || ''
      // 处理URL编码和可能的引号问题
      if (url) {
        // 解码URL
        url = decodeURIComponent(url)
        // 移除可能的引号
        url = url.replace(/^['"]|['"]$/g, '')
      }
      return url
    }
  },
  mounted() {
    document.title = '直接预览 - Vue文档预览器'
    document.body.style.overflow = 'hidden'
    
    console.log('DirectViewer mounted')
    console.log('原始URL参数:', this.$route.query.url)
    console.log('处理后的URL:', this.documentUrl)
    
    if (this.documentUrl) {
      this.loadDocumentFromUrl()
    } else {
      console.warn('没有提供documentUrl')
    }
  },
  beforeDestroy() {
    document.body.style.overflow = ''
  },
  methods: {
    async loadDocumentFromUrl() {
      if (!this.documentUrl) {
        this.error = '未提供文档URL'
        console.error('未提供文档URL')
        return
      }

      console.log('开始加载文档:', this.documentUrl)
      this.loading = true
      this.error = null

      try {
        console.log('正在下载文件...')
        const fileBuffer = await DocumentUtils.downloadFile(this.documentUrl)
        console.log('文件下载成功，大小:', fileBuffer.byteLength, 'bytes')
        
        const fileName = DocumentUtils.getFileNameFromUrl(this.documentUrl)
        console.log('解析的文件名:', fileName)
        
        this.documentFile = fileBuffer
        this.documentFileName = fileName
        this.documentLoaded = true
        
        console.log('文档加载完成')
      } catch (error) {
        console.error('文档加载失败:', error)
        this.handleDocumentError(error)
      } finally {
        this.loading = false
      }
    },

    handleDocumentLoaded(result) {
      console.log('直接预览文档加载成功:', result)
    },

    handleDocumentError(error) {
      this.error = error.message || '文档加载失败'
      this.loading = false
      this.documentLoaded = false
    },

    retry() {
      this.error = null
      this.loadDocumentFromUrl()
    },

    goBack() {
      if (window.history.length > 1) {
        window.history.back()
      } else {
        this.$router.push('/')
      }
    }
  }
}
</script>

<style scoped>
.direct-viewer {
  height: 100vh;
  position: relative;
  background: #2c3e50;
  overflow: hidden;
}

.document-section {
  height: 100%;
  position: relative;
}

.loading-section {
  display: flex;
  justify-content: center;
  align-items: center;
  height: 100%;
  text-align: center;
  color: white;
}

.loading-content {
  max-width: 500px;
  padding: var(--spacing-xl);
}

.loading-icon {
  font-size: 72px;
  margin-bottom: var(--spacing-lg);
}

.loading-content h2 {
  margin-bottom: var(--spacing-md);
  color: white;
}

.loading-content p {
  margin-bottom: var(--spacing-lg);
  color: rgba(255, 255, 255, 0.8);
  font-size: var(--font-size-lg);
}

.loading-progress {
  margin: var(--spacing-lg) 0;
}

.loading-spinner {
  width: 40px;
  height: 40px;
  border: 4px solid rgba(255, 255, 255, 0.3);
  border-top: 4px solid white;
  border-radius: 50%;
  animation: spin 1s linear infinite;
  margin: 0 auto var(--spacing-md);
}

.loading-text {
  color: rgba(255, 255, 255, 0.8);
}

.error-section {
  display: flex;
  justify-content: center;
  align-items: center;
  height: 100%;
  text-align: center;
  color: white;
}

.error-content {
  max-width: 500px;
  padding: var(--spacing-xl);
}

.error-icon {
  font-size: 64px;
  margin-bottom: var(--spacing-lg);
}

.error-content h2 {
  color: #ff6b6b;
  margin-bottom: var(--spacing-md);
}

.error-message {
  margin-bottom: var(--spacing-lg);
  color: rgba(255, 255, 255, 0.8);
  font-size: var(--font-size-lg);
  line-height: 1.6;
}

.error-actions {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: var(--spacing-md);
}

.error-actions .btn {
  color: white;
  border-color: rgba(255, 255, 255, 0.3);
}

.error-actions .btn.btn-primary {
  background: var(--primary-color);
}

.error-actions .btn.btn-primary:hover {
  background: var(--primary-dark);
}

.error-actions .btn.btn-secondary {
  background: rgba(255, 255, 255, 0.1);
}

.error-actions .btn.btn-secondary:hover {
  background: rgba(255, 255, 255, 0.2);
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

@media (max-width: 768px) {
  .loading-content,
  .error-content {
    padding: var(--spacing-lg);
    margin: var(--spacing-md);
  }
}
</style>