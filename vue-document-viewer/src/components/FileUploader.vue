<template>
  <div class="file-uploader">
    <div 
      class="file-uploader__area"
      :class="{ dragover: isDragOver }"
      @drop="handleDrop"
      @dragover="handleDragOver"
      @dragenter="handleDragEnter"
      @dragleave="handleDragLeave"
    >
      <input
        ref="fileInput"
        type="file"
        class="file-input"
        :accept="acceptedFormats"
        @change="handleFileSelect"
      >
      
      <div class="upload-icon">
        <svg width="64" height="64" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <polyline points="14,2 14,8 20,8" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <line x1="16" y1="13" x2="8" y2="13" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <line x1="16" y1="17" x2="8" y2="17" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <polyline points="10,9 9,9 8,9" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </div>
      
      <h3 class="upload-title">
        {{ isDragOver ? '释放文件开始预览' : '本地文件预览器' }}
      </h3>
      
      <p class="upload-description">
        {{ isDragOver ? '' : '选择本地文档文件进行在线预览，支持PDF、Word、PowerPoint、Excel等格式' }}
      </p>
      
      <div v-if="!isDragOver" class="supported-formats">
        <div class="format-group">
          <span class="format-badge pdf">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
              <path d="M8.267 14.68c-.184 0-.308.018-.372.036v1.178c.076.018.171.018.276.018.142 0 .217-.06.217-.158v-.853c0-.114-.083-.221-.121-.221z"/>
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
            </svg>
            PDF
          </span>
          <span class="format-badge docx">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
            </svg>
            Word
          </span>
          <span class="format-badge pptx">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
            </svg>
            PowerPoint
          </span>
          <span class="format-badge xlsx">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
            </svg>
            Excel
          </span>
        </div>
      </div>
      
      <div v-if="!isDragOver" class="upload-actions">
        <button class="btn btn-primary btn-upload"  @click="openFileDialog">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="7,10 12,15 17,10"/>
            <line x1="12" y1="15" x2="12" y2="3"/>
          </svg>
          选择文件
        </button>
        <p class="upload-hint">或拖拽文件到此区域</p>
        <p class="size-hint">支持最大 {{ maxSizeMB }}MB 的文件</p>
      </div>
    </div>

    <!-- 错误信息显示 -->
    <div v-if="error" class="upload-error">
      <div class="error-icon">⚠️</div>
      <div class="error-message">{{ error }}</div>
      <button @click="clearError" class="btn btn-sm btn-secondary">关闭</button>
    </div>

    <!-- 文件信息显示 -->
    <div v-if="selectedFile && !error" class="file-info">
      <div class="file-info__icon">{{ getFileIcon() }}</div>
      <div class="file-info__details">
        <div class="file-name">{{ selectedFile.name }}</div>
        <div class="file-meta">
          <span>{{ formatFileSize(selectedFile.size) }}</span>
          <span>{{ getFileType() }}</span>
        </div>
      </div>
      <div class="file-info__actions">
        <button @click="clearFile" class="btn btn-sm btn-secondary">移除</button>
        <button @click="previewFile" class="btn btn-sm btn-primary">预览</button>
      </div>
    </div>
  </div>
</template>

<script>
import { DocumentUtils } from '@/utils/documentUtils'

export default {
  name: 'FileUploader',
  props: {
    maxSize: {
      type: Number,
      default: 50 * 1024 * 1024 // 50MB
    },
    acceptedTypes: {
      type: Array,
      default: () => ['pdf', 'docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls', 'csv']
    },
    autoPreview: {
      type: Boolean,
      default: true
    }
  },
  data() {
    return {
      selectedFile: null,
      isDragOver: false,
      error: null,
      dragCounter: 0 // 用于处理嵌套拖拽事件
    }
  },
  computed: {
    maxSizeMB() {
      return Math.round(this.maxSize / (1024 * 1024))
    },
    acceptedFormats() {
      return this.acceptedTypes.map(type => `.${type}`).join(',')
    }
  },
  methods: {
    preventDefaults(event) {
      event.preventDefault()
      event.stopPropagation()
    },
    openFileDialog() {
      this.$refs.fileInput.click()
    },

    handleFileSelect(event) {
      const file = event.target.files[0]
      if (file) {
        this.processFile(file)
      }
    },

    handleDrop(event) {
      event.preventDefault()
      event.stopPropagation()
      
      this.isDragOver = false
      this.dragCounter = 0
      
      const files = event.dataTransfer.files
      if (files.length > 0) {
        this.processFile(files[0])
      }
    },

    handleDragOver(event) {
      event.preventDefault()
      event.stopPropagation()
    },

    handleDragEnter(event) {
      event.preventDefault()
      event.stopPropagation()
      
      this.dragCounter++
      this.isDragOver = true
    },

    handleDragLeave(event) {
      event.preventDefault()
      event.stopPropagation()
      
      this.dragCounter--
      if (this.dragCounter === 0) {
        this.isDragOver = false
      }
    },

    processFile(file) {
      this.clearError()
      
      // 验证文件
      const validation = this.validateFile(file)
      if (!validation.valid) {
        this.showError(validation.error)
        return
      }

      this.selectedFile = file
      this.$emit('file-selected', file)

      if (this.autoPreview) {
        this.previewFile()
      }
    },

    validateFile(file) {
      // 检查文件大小
      if (file.size > this.maxSize) {
        return {
          valid: false,
          error: `文件大小超出限制，最大支持 ${this.maxSizeMB}MB`
        }
      }

      // 检查文件类型
      const extension = DocumentUtils.getFileExtension(file.name)
      if (!this.acceptedTypes.includes(extension)) {
        return {
          valid: false,
          error: `不支持的文件格式：${extension}`
        }
      }

      if (!DocumentUtils.isSupportedFormat(extension)) {
        return {
          valid: false,
          error: `当前不支持此文件格式：${extension}`
        }
      }

      return { valid: true }
    },

    previewFile() {
      if (this.selectedFile) {
        this.$emit('preview-file', this.selectedFile)
      }
    },

    clearFile() {
      this.selectedFile = null
      this.$refs.fileInput.value = ''
      this.$emit('file-cleared')
    },

    showError(message) {
      this.error = message
      // 清除选中的文件
      this.selectedFile = null
      this.$refs.fileInput.value = ''
    },

    clearError() {
      this.error = null
    },

    // 辅助方法
    getFileIcon() {
      if (!this.selectedFile) return '📄'
      
      const extension = DocumentUtils.getFileExtension(this.selectedFile.name)
      const icons = {
        pdf: '📄',
        docx: '📝',
        doc: '📝',
        pptx: '📊',
        ppt: '📊',
        xlsx: '📈',
        xls: '📈',
        csv: '📋'
      }
      return icons[extension] || '📄'
    },

    getFileType() {
      if (!this.selectedFile) return ''
      
      const extension = DocumentUtils.getFileExtension(this.selectedFile.name)
      const types = {
        pdf: 'PDF文档',
        docx: 'Word文档',
        doc: 'Word文档',
        pptx: 'PowerPoint演示文稿',
        ppt: 'PowerPoint演示文稿',
        xlsx: 'Excel电子表格',
        xls: 'Excel电子表格',
        csv: 'CSV数据文件'
      }
      return types[extension] || '未知格式'
    },

    formatFileSize(bytes) {
      return DocumentUtils.formatFileSize(bytes)
    }
  },

  mounted() {
    // 防止整个页面的拖拽默认行为
    document.addEventListener('dragover', this.preventDefaults)
    document.addEventListener('drop', this.preventDefaults)
  },

  beforeDestroy() {
    document.removeEventListener('dragover', this.preventDefaults)
    document.removeEventListener('drop', this.preventDefaults)
  },
}
</script>

<style scoped>
.file-uploader {
  width: 100%;
  max-width: 800px;
  margin: 0 auto;
  padding: 2rem;
}

.file-uploader__area {
  border: 2px dashed #e1e5e9;
  border-radius: 16px;
  /* padding: 60px 32px; */
  padding: 10px 32px;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  background: linear-gradient(135deg, #fafbfc 0%, #f6f8fa 100%);
  /* min-height: 400px; */
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  position: relative;
  overflow: hidden;
}

.file-uploader__area::before {
  content: '';
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: linear-gradient(45deg, transparent 30%, rgba(255,255,255,0.1) 50%, transparent 70%);
  transform: translateX(-100%);
  transition: transform 0.6s;
}

.file-uploader__area:hover {
  border-color: #2563eb;
  background: linear-gradient(135deg, #f8faff 0%, #eff6ff 100%);
  box-shadow: 0 10px 25px rgba(37, 99, 235, 0.1);
  transform: translateY(-2px);
}

.file-uploader__area:hover::before {
  transform: translateX(100%);
}

.file-uploader__area.dragover {
  border-color: #2563eb;
  background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%);
  transform: scale(1.02);
  box-shadow: 0 20px 40px rgba(37, 99, 235, 0.2);
}

.file-input {
  display: none;
}

.upload-icon {
  color: #6b7280;
  margin-bottom: 1.5rem;
  transition: all 0.3s ease;
}

.file-uploader__area:hover .upload-icon {
  color: #2563eb;
  transform: scale(1.1);
}

.file-uploader__area.dragover .upload-icon {
  color: #1d4ed8;
  transform: scale(1.2);
}

.upload-title {
  font-size: 1.75rem;
  font-weight: 700;
  color: #1f2937;
  margin: 0 0 1rem 0;
  transition: all 0.3s ease;
}

.file-uploader__area.dragover .upload-title {
  color: #1d4ed8;
}

.upload-description {
  font-size: 1.1rem;
  color: #6b7280;
  margin: 0 0 2rem 0;
  line-height: 1.6;
  max-width: 500px;
}

.supported-formats {
  margin-bottom: 2rem;
}

.format-group {
  display: flex;
  flex-wrap: wrap;
  gap: 0.75rem;
  justify-content: center;
}

.format-badge {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.5rem 1rem;
  background: white;
  border: 1px solid #e5e7eb;
  border-radius: 8px;
  font-size: 0.875rem;
  font-weight: 500;
  color: #374151;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
  transition: all 0.2s ease;
}

.format-badge:hover {
  transform: translateY(-1px);
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}

.format-badge.pdf { border-left: 4px solid #dc2626; }
.format-badge.docx { border-left: 4px solid #2563eb; }
.format-badge.pptx { border-left: 4px solid #ea580c; }
.format-badge.xlsx { border-left: 4px solid #16a34a; }

.upload-actions {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 1rem;
}

.btn-upload {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.75rem 2rem;
  font-size: 1rem;
  font-weight: 600;
  border-radius: 8px;
  transition: all 0.2s ease;
}

.btn-primary {
  background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
  color: white;
  border: none;
  box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
}

.btn-primary:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 20px rgba(37, 99, 235, 0.4);
}

.upload-hint {
  font-size: 0.95rem;
  color: #9ca3af;
  margin: 0;
}

.size-hint {
  font-size: 0.85rem;
  color: #d1d5db;
  margin: 0;
}

.upload-error {
  background: #ef4444;
  border: 1px solid #d1d5db;
  border-radius: 8px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
  padding: 1rem;
  color: white;
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 1rem;
  max-width: 500px;
}

.upload-error .error-icon {
  font-size: 20px;
}

.upload-error .error-message {
  flex: 1;
}

.upload-error .btn {
  background: rgba(255, 255, 255, 0.2);
  color: white;
  border: 1px solid rgba(255, 255, 255, 0.3);
}

.upload-error .btn:hover {
  background: rgba(255, 255, 255, 0.3);
}

.file-info {
  background: white;
  border: 1px solid #d1d5db;
  border-radius: 8px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
  padding: 1rem;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 1rem;
  max-width: 500px;
  width: 100%;
}

.file-info__icon {
  font-size: 24px;
  flex-shrink: 0;
}

.file-info__details {
  flex: 1;
  min-width: 0;
}

.file-info__details .file-name {
  font-weight: 600;
  margin-bottom: 0.25rem;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.file-info__details .file-meta {
  font-size: var(--font-size-sm);
  color: var(--text-secondary);
  display: flex;
  gap: 0.5rem;
}

.file-info__actions {
  display: flex;
  gap: 0.5rem;
  flex-shrink: 0;
}

@media (max-width: 768px) {
  .file-info {
    flex-direction: column;
    text-align: center;
  }
  
  .file-info__actions {
    justify-content: center;
  }
}
</style>
