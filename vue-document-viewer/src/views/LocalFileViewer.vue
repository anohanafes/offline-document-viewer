<template>
  <div class="local-file-viewer">
    <!-- 欢迎区域 -->
    <div v-if="!selectedFile && !documentLoaded" class="welcome-section">
      <div class="welcome-content">
        <div class="welcome-icon">📄</div>
        <h1 class="welcome-title">本地文件预览器</h1>
        <p class="welcome-description">
          选择本地文档文件进行在线预览，支持PDF、Word、PowerPoint、Excel等格式
        </p>
      </div>
    </div>

    <!-- 文件上传区域 -->
    <div v-if="!documentLoaded" class="upload-section">
      <FileUploader
        :max-size="maxFileSize"
        :accepted-types="acceptedTypes"
        :auto-preview="false"
        @file-selected="handleFileSelected"
        @preview-file="handlePreviewFile"
        @file-cleared="handleFileCleared"
      />
    </div>

    <!-- 文档预览区域 -->
    <div v-if="documentLoaded" class="document-section">
      <DocumentViewer
        :file="selectedFile"
        :file-name="fileName"
        @document-loaded="handleDocumentLoaded"
        @document-error="handleDocumentError"
        @close="handleCloseDocument"
        ref="documentViewer"
      />
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
import FileUploader from '@/components/FileUploader.vue'

export default {
  name: 'LocalFileViewer',
  components: {
    DocumentViewer,
    FileUploader
  },
  data() {
    return {
      selectedFile: null,
      fileName: '',
      documentLoaded: false,
      error: null,
      
      maxFileSize: 50 * 1024 * 1024, // 50MB
      acceptedTypes: ['pdf', 'docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls', 'csv']
    }
  },
  methods: {
    handleFileSelected(file) {
      this.selectedFile = file
      this.fileName = file.name
      this.clearError()
    },

    handlePreviewFile(file) {
      this.selectedFile = file
      this.fileName = file.name
      this.documentLoaded = true
      this.clearError()
    },

    handleFileCleared() {
      this.selectedFile = null
      this.fileName = ''
      this.documentLoaded = false
      this.clearError()
    },

    handleDocumentLoaded(result) {
      console.log('文档加载成功:', result)
    },

    handleDocumentError(error) {
      this.error = error.message || '文档加载失败'
      this.documentLoaded = false
    },

    handleCloseDocument() {
      this.documentLoaded = false
      this.selectedFile = null
      this.fileName = ''
      this.clearError()
    },

    retry() {
      if (this.selectedFile) {
        this.clearError()
        this.documentLoaded = true
      }
    },

    clearError() {
      this.error = null
    }
  },

  mounted() {
    document.title = '本地文件预览 - Vue文档预览器'
  }
}
</script>

<style scoped>
.local-file-viewer {
  height: 100%;
  display: flex;
  flex-direction: column;
  position: relative;
}

.welcome-section {
  flex: 1;
  display: flex;
  justify-content: center;
  align-items: center;
  padding: var(--spacing-xl);
  background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
}

.welcome-content {
  text-align: center;
  max-width: 600px;
}

.welcome-icon {
  font-size: 72px;
  margin-bottom: var(--spacing-lg);
}

.welcome-title {
  font-size: var(--font-size-xxl);
  font-weight: 600;
  color: var(--text-primary);
  margin-bottom: var(--spacing-md);
}

.welcome-description {
  font-size: var(--font-size-lg);
  color: var(--text-secondary);
  line-height: 1.6;
  margin-bottom: var(--spacing-xl);
}

.upload-section {
  flex: 1;
}

.document-section {
  height: 100%;
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
  max-width: 400px;
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
}

.error-actions {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: var(--spacing-md);
}

@media (max-width: 768px) {
  .welcome-section {
    padding: var(--spacing-lg);
  }
  
  .welcome-icon {
    font-size: 56px;
  }
  
  .welcome-title {
    font-size: var(--font-size-xl);
  }
  
  .welcome-description {
    font-size: var(--font-size-md);
  }
  
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