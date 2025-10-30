<template>
  <div class="url-viewer">
    <!-- URL输入区域 -->
    <div v-if="!documentLoaded" class="input-section">
      <UrlInput
        :initial-url="initialUrl"
        :auto-load="!!initialUrl"
        @document-loaded="handleDocumentLoaded"
        @document-error="handleDocumentError"
      />
    </div>

    <!-- 文档预览区域 -->
    <div v-if="documentLoaded" class="document-section">
      <DocumentViewer
        :file="documentFile"
        :file-name="documentFileName"
        @document-loaded="handleViewerDocumentLoaded"
        @document-error="handleViewerDocumentError"
        @close="handleCloseDocument"
        ref="documentViewer"
      />
    </div>

    <!-- 加载状态 -->
    <div v-if="loading" class="loading-section">
      <div class="loading-content">
        <div class="loading-spinner"></div>
        <div class="loading-text">{{ loadingText }}</div>
        <button @click="cancelLoading" class="btn btn-secondary">取消</button>
      </div>
    </div>

    <!-- 错误提示 -->
    <div v-if="error" class="error-section">
      <div class="error-content">
        <div class="error-icon">❌</div>
        <h3 class="error-title">文档加载失败</h3>
        <p class="error-message">{{ error }}</p>
        <div class="error-actions">
          <button @click="clearError" class="btn btn-secondary">关闭</button>
          <button @click="retry" class="btn btn-primary">重试</button>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import DocumentViewer from '@/components/DocumentViewer.vue'
import UrlInput from '@/components/UrlInput.vue'
import { DocumentUtils } from '@/utils/documentUtils'

export default {
  name: 'UrlViewer',
  components: {
    DocumentViewer,
    UrlInput
  },
  data() {
    return {
      documentLoaded: false,
      documentFile: null,
      documentFileName: '',
      documentUrl: '',
      
      loading: false,
      loadingText: '正在下载文档...',
      
      error: null,
      errorDetails: null
    }
  },
  computed: {
    initialUrl() {
      return this.$route.query.url || ''
    }
  },
  mounted() {
    document.title = 'URL文档预览 - Vue文档预览器'
  },
  methods: {
    async handleDocumentLoaded(data) {
      this.loading = true
      this.error = null
      this.loadingText = '正在处理文档...'
      
      try {
        this.documentFile = data.file
        this.documentFileName = data.fileName
        this.documentUrl = data.url
        
        setTimeout(() => {
          this.documentLoaded = true
          this.loading = false
        }, 500)
        
      } catch (error) {
        this.handleDocumentError(error)
      }
    },

    handleDocumentError(error) {
      this.error = error.message || '文档加载失败'
      this.loading = false
      this.documentLoaded = false
    },

    handleViewerDocumentLoaded(result) {
      console.log('文档查看器加载成功:', result)
    },

    handleViewerDocumentError(error) {
      this.handleDocumentError(error)
    },

    handleCloseDocument() {
      this.documentLoaded = false
      this.documentFile = null
      this.documentFileName = ''
      this.documentUrl = ''
      this.clearError()
      
      if (this.$route.query.url) {
        this.$router.replace({ query: {} })
      }
    },

    cancelLoading() {
      this.loading = false
    },

    retry() {
      if (this.documentUrl) {
        this.clearError()
        window.location.reload()
      }
    },

    clearError() {
      this.error = null
      this.errorDetails = null
    }
  }
}
</script>

<style scoped>
.url-viewer {
  height: 100%;
  display: flex;
  flex-direction: column;
  position: relative;
}

.input-section {
  flex: 1;
  display: flex;
  flex-direction: column;
}

.document-section {
  height: 100%;
}

.loading-section {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  display: flex;
  justify-content: center;
  align-items: center;
  background: rgba(255, 255, 255, 0.95);
  z-index: 100;
}

.loading-content {
  text-align: center;
  max-width: 300px;
}

.loading-spinner {
  width: 60px;
  height: 60px;
  border: 6px solid var(--border-color);
  border-top: 6px solid var(--primary-color);
  border-radius: 50%;
  animation: spin 1s linear infinite;
  margin: 0 auto var(--spacing-md);
}

.loading-text {
  font-size: var(--font-size-lg);
  color: var(--text-primary);
  margin-bottom: var(--spacing-md);
}

.error-section {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  z-index: 100;
}

.error-content {
  background: var(--background-primary);
  border: 1px solid var(--border-color);
  border-radius: var(--border-radius-md);
  box-shadow: var(--shadow-light);
  padding: var(--spacing-xl);
  text-align: center;
  max-width: 500px;
}

.error-icon {
  font-size: 48px;
  margin-bottom: var(--spacing-md);
}

.error-title {
  color: var(--error-color);
  margin-bottom: var(--spacing-md);
}

.error-message {
  color: var(--text-secondary);
  margin-bottom: var(--spacing-lg);
  line-height: 1.6;
}

.error-actions {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: var(--spacing-md);
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

@media (max-width: 768px) {
  .error-content {
    margin: var(--spacing-md);
    max-width: calc(100vw - 2 * var(--spacing-md));
  }
  
  .error-actions {
    flex-direction: column;
    align-items: stretch;
  }
}
</style>