<template>
  <div class="viewer-controls">
    <!-- PDF 控制 -->
    <div v-if="documentType === 'pdf'" class="controls-group pdf-controls">
      <div class="control-section">
        <button 
          @click="$emit('prev-page')" 
          :disabled="currentPage <= 1"
          class="btn btn-sm btn-secondary"
          title="上一页 (←)"
        >
          ◀
        </button>
        <span class="page-info">
          {{ currentPage }} / {{ totalPages }}
        </span>
        <button 
          @click="$emit('next-page')" 
          :disabled="currentPage >= totalPages"
          class="btn btn-sm btn-secondary"
          title="下一页 (→)"
        >
          ▶
        </button>
      </div>
      
      <div class="control-section">
        <input
          v-model.number="pageInput"
          @keydown.enter="goToPage"
          @blur="goToPage"
          type="number"
          :min="1"
          :max="totalPages"
          class="page-input"
          :placeholder="`1-${totalPages}`"
        >
        <button @click="goToPage" class="btn btn-sm btn-secondary">跳转</button>
      </div>
    </div>

    <!-- PPTX 控制 -->
    <div v-if="documentType === 'pptx'" class="controls-group pptx-controls">
      <div class="control-section">
        <button 
          @click="$emit('prev-slide')" 
          :disabled="currentSlide <= 1"
          class="btn btn-sm btn-secondary"
          title="上一张 (←)"
        >
          ◀
        </button>
        <span class="slide-info">
          第 {{ currentSlide }} 张 / 共 {{ totalSlides }} 张
        </span>
        <button 
          @click="$emit('next-slide')" 
          :disabled="currentSlide >= totalSlides"
          class="btn btn-sm btn-secondary"
          title="下一张 (→)"
        >
          ▶
        </button>
      </div>
      
      <div class="control-section">
        <button 
          @click="$emit('toggle-slideshow')" 
          class="btn btn-sm btn-primary"
          title="幻灯片播放"
        >
          {{ isSlideshow ? '⏸️ 停止播放' : '▶️ 开始播放' }}
        </button>
      </div>
    </div>

    <!-- Excel 工作表控制 -->
    <div v-if="documentType === 'xlsx' && worksheets?.length > 1" class="controls-group xlsx-controls">
      <div class="control-section">
        <label class="control-label">工作表:</label>
        <select 
          :value="currentWorksheet" 
          @change="$emit('worksheet-change', $event.target.value)"
          class="worksheet-select"
        >
          <option 
            v-for="(sheet, index) in worksheets" 
            :key="index" 
            :value="index"
          >
            {{ sheet.name }}
          </option>
        </select>
      </div>
      
      <div class="control-section">
        <span class="worksheet-info">
          {{ getCurrentWorksheetInfo() }}
        </span>
      </div>
    </div>

    <!-- 缩放控制 (通用) -->
    <div v-if="supportsZoom" class="controls-group zoom-controls">
      <div class="control-section">
        <button 
          @click="$emit('zoom-out')" 
          :disabled="zoomLevel <= minZoom"
          class="btn btn-sm btn-secondary"
          title="缩小 (Ctrl + -)"
        >
          🔍-
        </button>
        <span class="zoom-level">{{ Math.round(zoomLevel * 100) }}%</span>
        <button 
          @click="$emit('zoom-in')" 
          :disabled="zoomLevel >= maxZoom"
          class="btn btn-sm btn-secondary"
          title="放大 (Ctrl + +)"
        >
          🔍+
        </button>
      </div>
      
      <div class="control-section">
        <button 
          @click="$emit('fit-to-width')" 
          class="btn btn-sm btn-secondary"
          title="适应宽度"
        >
          适应宽度
        </button>
        <button 
          @click="$emit('actual-size')" 
          class="btn btn-sm btn-secondary"
          title="实际大小"
        >
          实际大小
        </button>
      </div>
    </div>

    <!-- 文档操作 -->
    <div class="controls-group document-actions">
      <div class="control-section">
        <button 
          @click="$emit('download')" 
          class="btn btn-sm btn-outline"
          title="下载文档"
        >
          📥 下载
        </button>
        <button 
          @click="$emit('print')" 
          class="btn btn-sm btn-outline"
          title="打印文档"
        >
          🖨️ 打印
        </button>
        <button 
          @click="$emit('fullscreen')" 
          class="btn btn-sm btn-outline"
          title="全屏显示"
        >
          {{ isFullscreen ? '📴 退出全屏' : '📺 全屏' }}
        </button>
      </div>
    </div>

    <!-- 更多选项 -->
    <div class="controls-group more-options">
      <div class="control-section">
        <button 
          @click="showMoreOptions = !showMoreOptions"
          class="btn btn-sm btn-secondary"
          title="更多选项"
        >
          ⚙️ 更多
        </button>
      </div>
    </div>

    <!-- 展开的更多选项 -->
    <div v-if="showMoreOptions" class="more-options-panel">
      <div class="option-item">
        <label>
          <input 
            type="checkbox" 
            :checked="autoFit" 
            @change="$emit('auto-fit-change', $event.target.checked)"
          >
          自动适应
        </label>
      </div>
      
      <div v-if="documentType === 'pdf'" class="option-item">
        <label>
          <input 
            type="checkbox" 
            :checked="continuousView" 
            @change="$emit('continuous-view-change', $event.target.checked)"
          >
          连续浏览
        </label>
      </div>
      
      <div class="option-item">
        <label>
          背景色:
          <input 
            type="color" 
            :value="backgroundColor" 
            @change="$emit('background-color-change', $event.target.value)"
            class="color-input"
          >
        </label>
      </div>
    </div>
  </div>
</template>

<script>
export default {
  name: 'ViewerControls',
  props: {
    documentType: {
      type: String,
      required: true
    },
    
    // PDF 相关
    currentPage: {
      type: Number,
      default: 1
    },
    totalPages: {
      type: Number,
      default: 0
    },
    
    // PPTX 相关
    currentSlide: {
      type: Number,
      default: 1
    },
    totalSlides: {
      type: Number,
      default: 0
    },
    isSlideshow: {
      type: Boolean,
      default: false
    },
    
    // Excel 相关
    worksheets: {
      type: Array,
      default: () => []
    },
    currentWorksheet: {
      type: Number,
      default: 0
    },
    
    // 缩放相关
    zoomLevel: {
      type: Number,
      default: 1
    },
    minZoom: {
      type: Number,
      default: 0.25
    },
    maxZoom: {
      type: Number,
      default: 3
    },
    supportsZoom: {
      type: Boolean,
      default: true
    },
    
    // 其他状态
    isFullscreen: {
      type: Boolean,
      default: false
    },
    autoFit: {
      type: Boolean,
      default: false
    },
    continuousView: {
      type: Boolean,
      default: false
    },
    backgroundColor: {
      type: String,
      default: '#f5f5f5'
    }
  },
  
  data() {
    return {
      pageInput: this.currentPage,
      showMoreOptions: false
    }
  },
  
  watch: {
    currentPage(newPage) {
      this.pageInput = newPage
    }
  },
  
  methods: {
    goToPage() {
      const page = parseInt(this.pageInput)
      if (page >= 1 && page <= this.totalPages && page !== this.currentPage) {
        this.$emit('go-to-page', page)
      } else {
        // 重置为当前页
        this.pageInput = this.currentPage
      }
    },
    
    getCurrentWorksheetInfo() {
      if (!this.worksheets || !this.worksheets[this.currentWorksheet]) {
        return ''
      }
      
      const sheet = this.worksheets[this.currentWorksheet]
      return `${sheet.rowCount} 行 × ${sheet.colCount} 列`
    }
  }
}
</script>

<style scoped>
.viewer-controls {
  background: var(--background-primary);
  border-bottom: 1px solid var(--border-color);
  padding: var(--spacing-sm) var(--spacing-md);
  display: flex;
  align-items: center;
  gap: var(--spacing-lg);
  flex-wrap: wrap;
  position: relative;
}

.viewer-controls .controls-group {
  display: flex;
  align-items: center;
  gap: var(--spacing-md);
  padding: var(--spacing-xs) 0;
}

.viewer-controls .controls-group .control-section {
  display: flex;
  align-items: center;
  gap: var(--spacing-sm);
}

.viewer-controls .control-label {
  font-size: var(--font-size-sm);
  color: var(--text-secondary);
  white-space: nowrap;
}

.viewer-controls .page-info,
.viewer-controls .slide-info {
  font-size: var(--font-size-sm);
  color: var(--text-secondary);
  min-width: 80px;
  text-align: center;
  white-space: nowrap;
}

.viewer-controls .zoom-level {
  font-size: var(--font-size-sm);
  color: var(--text-secondary);
  min-width: 50px;
  text-align: center;
  font-weight: 500;
}

.viewer-controls .page-input {
  width: 60px;
  padding: 4px 6px;
  border: 1px solid var(--border-color);
  border-radius: var(--border-radius-sm);
  font-size: var(--font-size-sm);
  text-align: center;
}

.viewer-controls .page-input:focus {
  outline: none;
  border-color: var(--primary-color);
}

.viewer-controls .worksheet-select {
  padding: 4px 8px;
  border: 1px solid var(--border-color);
  border-radius: var(--border-radius-sm);
  font-size: var(--font-size-sm);
  background: var(--background-primary);
}

.viewer-controls .worksheet-select:focus {
  outline: none;
  border-color: var(--primary-color);
}

.viewer-controls .worksheet-info {
  font-size: var(--font-size-xs);
  color: var(--text-secondary);
  white-space: nowrap;
}

/* 分隔线 */
.viewer-controls .controls-group:not(:last-child)::after {
  content: '';
  width: 1px;
  height: 20px;
  background: var(--border-color);
  margin-left: var(--spacing-md);
}

.viewer-controls .more-options-panel {
  position: absolute;
  top: 100%;
  right: var(--spacing-md);
  background: var(--background-primary);
  border: 1px solid var(--border-color);
  border-radius: var(--border-radius-md);
  box-shadow: var(--shadow-medium);
  padding: var(--spacing-md);
  z-index: 10;
  min-width: 200px;
}

.viewer-controls .more-options-panel .option-item {
  margin-bottom: var(--spacing-sm);
}

.viewer-controls .more-options-panel .option-item:last-child {
  margin-bottom: 0;
}

.viewer-controls .more-options-panel .option-item label {
  display: flex;
  align-items: center;
  gap: var(--spacing-sm);
  font-size: var(--font-size-sm);
  cursor: pointer;
}

.viewer-controls .more-options-panel .option-item label input[type="checkbox"] {
  margin: 0;
}

.viewer-controls .more-options-panel .option-item label .color-input {
  width: 30px;
  height: 20px;
  border: none;
  border-radius: var(--border-radius-sm);
  cursor: pointer;
}

/* 响应式设计 */
@media (max-width: 1024px) {
  .viewer-controls {
    gap: var(--spacing-md);
  }
  
  .viewer-controls .controls-group {
    gap: var(--spacing-sm);
  }
}

@media (max-width: 768px) {
  .viewer-controls {
    flex-direction: column;
    align-items: stretch;
    gap: var(--spacing-sm);
  }
  
  .viewer-controls .controls-group {
    justify-content: center;
    flex-wrap: wrap;
  }
  
  .viewer-controls .controls-group::after {
    display: none;
  }
  
  .viewer-controls .more-options-panel {
    position: relative;
    top: auto;
    right: auto;
    margin-top: var(--spacing-sm);
  }
}

@media (max-width: 576px) {
  .viewer-controls {
    padding: var(--spacing-sm);
  }
  
  .viewer-controls .controls-group .control-section {
    flex-wrap: wrap;
    justify-content: center;
  }
  
  .viewer-controls .page-info,
  .viewer-controls .slide-info {
    min-width: auto;
  }
}
</style>
