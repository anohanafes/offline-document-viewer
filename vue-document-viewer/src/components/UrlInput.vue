<template>
  <div class="url-input">
    <div class="url-input-form">
      <h2>URL文档预览</h2>
      <p class="form-description">
        输入文档的在线地址，支持直接访问的PDF、DOCX、PPTX等文件链接
      </p>

      <div class="input-field">
        <label for="document-url">文档URL地址</label>
        <input
          id="document-url"
          v-model="url"
          type="url"
          placeholder="https://example.com/document.pdf"
          :class="{ error: urlError }"
          @input="clearError"
          @keydown.enter="loadDocument"
        >
        <div v-if="urlError" class="field-error">{{ urlError }}</div>
        <div class="field-hint">
          请确保文档链接可直接访问，且服务器支持跨域请求(CORS)
        </div>
      </div>

      <div class="url-actions">
        <button 
          @click="clearUrl" 
          :disabled="!url"
          class="btn btn-secondary"
        >
          清空
        </button>
        <button 
          @click="loadDocument" 
          :disabled="!url || loading"
          class="btn btn-primary"
        >
          <span v-if="loading">正在加载...</span>
          <span v-else>加载预览</span>
        </button>
      </div>

      <!-- 加载进度 -->
      <div v-if="loading" class="loading-progress">
        <div class="progress-bar">
          <div 
            class="progress-fill" 
            :style="{ width: `${loadingProgress}%` }"
          ></div>
        </div>
        <div class="progress-text">{{ loadingText }}</div>
      </div>

      <!-- 错误信息 -->
      <div v-if="error" class="error-message">
        <div class="error-icon">❌</div>
        <div class="error-content">
          <div class="error-title">加载失败</div>
          <div class="error-details">{{ error }}</div>
        </div>
        <button @click="clearError" class="btn btn-sm btn-secondary">关闭</button>
      </div>
    </div>
  </div>
</template>

<script>
import { DocumentUtils } from '@/utils/documentUtils'

export default {
  name: 'UrlInput',
  props: {
    initialUrl: {
      type: String,
      default: ''
    },
    autoLoad: {
      type: Boolean,
      default: false
    }
  },
  data() {
    return {
      url: this.initialUrl,
      urlError: null,
      loading: false,
      loadingProgress: 0,
      loadingText: '正在下载文档...',
      error: null
    }
  },
  watch: {
    initialUrl: {
      handler(newUrl) {
        this.url = newUrl
        if (newUrl && this.autoLoad) {
          this.loadDocument()
        }
      },
      immediate: true
    }
  },
  methods: {
    validateUrl() {
      if (!this.url) {
        this.urlError = '请输入文档URL地址'
        return false
      }

      try {
        const urlObj = new URL(this.url)
        if (!['http:', 'https:'].includes(urlObj.protocol)) {
          this.urlError = '请输入有效的HTTP或HTTPS地址'
          return false
        }
      } catch (error) {
        this.urlError = '请输入有效的URL地址'
        return false
      }

      const fileName = DocumentUtils.getFileNameFromUrl(this.url)
      const extension = DocumentUtils.getFileExtension(fileName)
      
      if (!DocumentUtils.isSupportedFormat(extension)) {
        this.urlError = `不支持的文件格式: ${extension}。支持的格式: PDF, DOCX, PPTX, XLSX, CSV`
        return false
      }

      return true
    },

    async loadDocument() {
      if (!this.validateUrl()) {
        return
      }

      this.loading = true
      this.loadingProgress = 0
      this.error = null
      this.loadingText = '正在下载文档...'

      try {
        const progressInterval = setInterval(() => {
          if (this.loadingProgress < 90) {
            this.loadingProgress += Math.random() * 20
          }
        }, 200)

        const fileBuffer = await DocumentUtils.downloadFile(this.url)
        
        clearInterval(progressInterval)
        this.loadingProgress = 100
        this.loadingText = '下载完成，正在处理...'

        const fileName = DocumentUtils.getFileNameFromUrl(this.url)
        
        this.$emit('document-loaded', {
          file: fileBuffer,
          fileName: fileName,
          url: this.url
        })
        
      } catch (error) {
        this.error = this.getErrorMessage(error)
        this.$emit('document-error', error)
      } finally {
        this.loading = false
        this.loadingProgress = 0
      }
    },

    getErrorMessage(error) {
      if (error.message.includes('CORS')) {
        return '跨域请求被阻止。请确保目标服务器支持CORS，或使用代理服务器。'
      }
      if (error.message.includes('404')) {
        return '文档不存在或地址无效 (404)'
      }
      if (error.message.includes('403')) {
        return '没有访问权限 (403)'
      }
      return error.message || '下载文档时发生未知错误'
    },

    clearUrl() {
      this.url = ''
      this.clearError()
    },

    clearError() {
      this.urlError = null
      this.error = null
    }
  }
}
</script>

<style scoped>
.url-input {
  padding: var(--spacing-lg);
  display: flex;
  justify-content: center;
  align-items: center;
  min-height: 100%;
}

.url-input-form {
  background: var(--background-primary);
  border: 1px solid var(--border-color);
  border-radius: var(--border-radius-md);
  box-shadow: var(--shadow-light);
  padding: var(--spacing-xl);
  max-width: 600px;
  width: 100%;
}

.url-input-form h2 {
  margin-bottom: var(--spacing-md);
  color: var(--text-primary);
}

.form-description {
  color: var(--text-secondary);
  margin-bottom: var(--spacing-lg);
  line-height: 1.6;
}

.input-field {
  margin-bottom: var(--spacing-lg);
}

.input-field label {
  display: block;
  margin-bottom: var(--spacing-sm);
  font-weight: 500;
  color: var(--text-primary);
}

.input-field input {
  width: 100%;
  padding: var(--spacing-sm) var(--spacing-md);
  border: 2px solid var(--border-color);
  border-radius: var(--border-radius-sm);
  font-size: var(--font-size-md);
  transition: all var(--transition-normal);
}

.input-field input:focus {
  outline: none;
  border-color: var(--primary-color);
  box-shadow: 0 0 0 3px rgba(33, 150, 243, 0.1);
}

.input-field input.error {
  border-color: var(--error-color);
}

.field-error {
  color: var(--error-color);
  font-size: var(--font-size-sm);
  margin-top: var(--spacing-xs);
}

.field-hint {
  color: var(--text-secondary);
  font-size: var(--font-size-sm);
  margin-top: var(--spacing-xs);
  line-height: 1.4;
}

.url-actions {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: var(--spacing-md);
  margin-bottom: var(--spacing-lg);
}

.loading-progress {
  margin-bottom: var(--spacing-lg);
}

.progress-bar {
  width: 100%;
  height: 6px;
  background: var(--background-secondary);
  border-radius: 3px;
  overflow: hidden;
  margin-bottom: var(--spacing-sm);
}

.progress-fill {
  height: 100%;
  background: var(--primary-color);
  transition: width 0.3s ease;
  border-radius: 3px;
}

.progress-text {
  text-align: center;
  font-size: var(--font-size-sm);
  color: var(--text-secondary);
}

.error-message {
  background: var(--error-color);
  color: white;
  border-radius: var(--border-radius-md);
  padding: var(--spacing-md);
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: var(--spacing-md);
}

.error-icon {
  font-size: 20px;
  flex-shrink: 0;
}

.error-content {
  flex: 1;
}

.error-title {
  font-weight: 600;
  margin-bottom: var(--spacing-xs);
}

.error-details {
  font-size: var(--font-size-sm);
  opacity: 0.9;
}

@media (max-width: 768px) {
  .url-input {
    padding: var(--spacing-md);
  }
  
  .url-input-form {
    padding: var(--spacing-lg);
  }
  
  .url-actions {
    flex-direction: column;
    align-items: stretch;
  }
  
  .error-message {
    flex-direction: column;
    text-align: center;
  }
}
</style>