<template>
  <div class="document-viewer">
    <!-- 文档显示区域 -->
    <div class="document-viewer__content">
      <!-- 加载状态 -->
      <div v-if="loading" class="document-viewer__loading">
        <div class="loading-spinner"></div>
        <div class="loading-text">{{ loadingText }}</div>
      </div>

      <!-- 错误状态 -->
      <div v-else-if="error" class="document-viewer__error">
        <div class="error-icon">❌</div>
        <h3 class="error-title">文档加载失败</h3>
        <p class="error-message">{{ error }}</p>
        <div class="error-actions">
          <button @click="retry" class="btn btn-primary">重试</button>
        </div>
      </div>

      <!-- PDF预览 -->
      <div v-show="fileInfo && fileInfo.type === 'pdf'" class="pdf-container">
        <div class="pdf-controls">
          <button @click="prevPage" :disabled="currentPage <= 1" class="btn btn-sm">上一页</button>
          <span>第 {{ currentPage }} 页 / 共 {{ totalPages }} 页</span>
          <button @click="nextPage" :disabled="currentPage >= totalPages" class="btn btn-sm">下一页</button>
        </div>
        <canvas ref="pdfCanvas" class="pdf-canvas"></canvas>
      </div>

      <!-- DOCX预览 -->
      <div v-show="fileInfo && fileInfo.type === 'docx'" class="docx-container">
        <div class="docx-content" v-html="docxContent"></div>
      </div>

      <!-- PPTX预览 -->
      <div v-show="fileInfo && fileInfo.type === 'pptx'" class="pptx-container">
        <div ref="pptxContainer" class="pptx-content"></div>
      </div>

      <!-- Excel预览 -->
      <div v-show="fileInfo && ['xlsx', 'xls', 'csv'].includes(fileInfo.type)" class="excel-viewer">
        <!-- 工作表标签 -->
        <div class="sheet-tabs">
          <button 
            v-for="(name, index) in sheetNames" 
            :key="index"
            :class="['sheet-tab', { active: currentSheetIndex === index }]"
            @click="switchSheet(index)"
          >
            {{ name || `Sheet${index + 1}` }}
          </button>
        </div>

        <!-- Excel内容 -->
        <div class="excel-content">
          <div class="excel-toolbar">
            <div class="sheet-info">
              <span>共 {{ sheetNames.length }} 个工作表</span>
            </div>
            <div class="view-controls">
              <button @click="exportCSV" class="btn btn-sm">📄 导出CSV</button>
              <button @click="exportJSON" class="btn btn-sm">📋 导出JSON</button>
            </div>
          </div>
          <div class="excel-table-container">
            <table id="excel-table" v-if="currentSheetData">
              <thead>
                <tr>
                  <th v-for="(header, index) in currentSheetData.headers" :key="index">
                    {{ header }}
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="(row, rowIndex) in currentSheetData.rows" :key="rowIndex">
                  <td v-for="(cell, colIndex) in row" :key="colIndex">
                    {{ cell }}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script>

export default {
  name: "DocumentViewer",
  props: {
    file: {
      type: [File, ArrayBuffer],
      default: null,
    },
    fileName: {
      type: String,
      default: "",
    },
    hideHeader: {
      type: Boolean,
      default: false,
    },
  },

  data() {
    return {
      loading: false,
      loadingText: "正在加载文档...",
      error: null,
      fileInfo: null,
      
      // PDF相关数据
      pdfDoc: null,
      currentPage: 1,
      totalPages: 0,
      
      // DOCX相关数据
      docxContent: '',
      
      // Excel相关数据
      workbook: null,
      sheetNames: [],
      currentSheetIndex: 0,
      currentSheetData: null,
    };
  },

  watch: {
    file: {
      immediate: true,
      handler(newFile) {
        if (newFile) {
          this.processFile();
        } else {
          // 重置状态
          this.loading = false;
          this.error = null;
          this.fileInfo = null;
          this.pdfDoc = null;
          this.currentPage = 1;
          this.totalPages = 0;
          this.docxContent = '';
          this.workbook = null;
          this.sheetNames = [];
          this.currentSheetIndex = 0;
          this.currentSheetData = null;
        }
      },
    },
  },

  methods: {
    async processFile() {
      if (!this.file) return;

      this.loading = true;
      this.error = null;
      this.loadingText = "正在加载文档...";

      try {
        // 重置所有状态
        this.pdfDoc = null;
        this.currentPage = 1;
        this.totalPages = 0;
        this.docxContent = '';
        this.workbook = null;
        this.sheetNames = [];
        this.currentSheetIndex = 0;
        this.currentSheetData = null;

        // 获取文件信息
        let fileName, fileSize;
        if (this.file instanceof File) {
          fileName = this.file.name;
          fileSize = this.formatFileSize(this.file.size);
        } else {
          fileName = this.fileName || "document";
          fileSize = this.formatFileSize(this.file.byteLength || this.file.length || 0);
        }

        const extension = this.getFileExtension(fileName);

        this.fileInfo = {
          fileName,
          fileSize,
          type: extension,
        };

        // 根据文件类型渲染
        await this.renderDocument(extension);
      } catch (error) {
        console.error("文档处理失败:", error);
        this.error = error.message;
      } finally {
        this.loading = false;
      }
    },

    async renderDocument(extension) {
      // 先设置fileInfo，触发Vue重新渲染对应的容器
      // fileInfo已在processFile中设置，这里再次等待DOM更新
      await this.$nextTick();
      
      // 再等一次确保条件渲染完成
      await this.$nextTick();

      switch (extension) {
        case "pdf":
          await this.renderPDF();
          break;
        case "docx":
          await this.renderDOCX();
          break;
        case "pptx":
          await this.renderPPTX();
          break;
        case "xlsx":
        case "xls":
        case "csv":
          await this.renderExcel();
          break;
        default:
          throw new Error(`不支持的文件格式: .${extension}`);
      }
    },

    async renderPDF() {
      try {
        console.log("开始渲染PDF");
        
        // 确保PDF.js已加载
        if (typeof window.pdfjsLib === "undefined") {
          throw new Error("PDF.js库未加载");
        }

        // 确保worker配置正确
        if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
          window.pdfjsLib.GlobalWorkerOptions.workerSrc = '/lib/pdf.worker.js';
          console.log("设置PDF.js worker路径:", window.pdfjsLib.GlobalWorkerOptions.workerSrc);
        }

        const arrayBuffer = this.file instanceof File ? await this.file.arrayBuffer() : this.file;
        console.log("PDF文件大小:", arrayBuffer.byteLength, "bytes");

        // 加载PDF文档
        const loadingTask = window.pdfjsLib.getDocument({
          data: arrayBuffer,
          verbosity: 0 // 减少控制台输出
        });
        
        this.pdfDoc = await loadingTask.promise;
        this.totalPages = this.pdfDoc.numPages;
        this.currentPage = 1;
        
        console.log("PDF文档加载完成，总页数:", this.totalPages);

        // 强制Vue重新渲染，确保PDF容器和canvas被创建
        this.$forceUpdate();
        await this.$nextTick();
        await this.$nextTick();
        
        // 等待canvas元素真正创建完成
        await new Promise(resolve => setTimeout(resolve, 200));
        
        // 渲染第一页
        await this.renderPDFPage(1);
      } catch (error) {
        console.error("PDF渲染失败:", error);
        throw error;
      }
    },

    async renderPDFPage(pageNum) {
      if (!this.pdfDoc) {
        console.error('PDF文档未加载');
        return;
      }

      try {
        console.log(`开始渲染PDF第${pageNum}页`);
        
        // 简单等待canvas元素，现在应该始终存在
        await this.$nextTick();
        await this.$nextTick();
        
        const canvas = this.$refs.pdfCanvas;
        if (!canvas) {
          console.error('Canvas元素仍然未找到');
          return;
        }

        console.log('✅ 找到canvas元素');
        
        const page = await this.pdfDoc.getPage(pageNum);
        const context = canvas.getContext("2d");

        // 清除之前的内容
        context.clearRect(0, 0, canvas.width, canvas.height);

        const viewport = page.getViewport({ scale: 1.2 });
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        const renderContext = {
          canvasContext: context,
          viewport: viewport,
        };

        await page.render(renderContext).promise;
        this.currentPage = pageNum;
        
        console.log(`✅ PDF第${pageNum}页渲染完成 (${canvas.width}x${canvas.height})`);
      } catch (error) {
        console.error(`PDF第${pageNum}页渲染失败:`, error);
        // 显示错误信息给用户
        const canvas = this.$refs.pdfCanvas;
        if (canvas) {
          const context = canvas.getContext("2d");
          context.fillStyle = "red";
          context.font = "16px Arial";
          context.fillText(`渲染失败: ${error.message}`, 10, 30);
        }
      }
    },

    async nextPage() {
      if (this.currentPage < this.totalPages) {
        await this.renderPDFPage(this.currentPage + 1);
      }
    },

    async prevPage() {
      if (this.currentPage > 1) {
        await this.renderPDFPage(this.currentPage - 1);
      }
    },

    async renderDOCX() {
      // 检查Mammoth是否已加载
      if (typeof window.mammoth === "undefined") {
        throw new Error("Mammoth库未加载");
      }

      // 使用Mammoth转换DOCX为HTML
      const arrayBuffer =
        this.file instanceof File ? await this.file.arrayBuffer() : this.file;

      const result = await window.mammoth.convertToHtml({ arrayBuffer });
      this.docxContent = result.value;

      // 如果有警告，在控制台显示
      if (result.messages && result.messages.length > 0) {
        console.warn("DOCX转换警告:", result.messages);
      }

      console.log("✅ DOCX预览渲染完成");
    },

    async renderPPTX() {
      // 检查JSZip和OfficeViewer是否已加载
      if (typeof window.JSZip === "undefined") {
        throw new Error("JSZip库未加载");
      }

      if (typeof window.OfficeViewer === "undefined") {
        throw new Error("OfficeViewer库未加载");
      }

      // 强制触发Vue重新渲染，确保PPTX容器被创建
      this.$forceUpdate();
      await this.$nextTick();
      await this.$nextTick();
      await this.$nextTick();
      
      // 多次检查确保容器已挂载
      let attempts = 0;
      let container = this.$refs.pptxContainer;
      
      while (!container && attempts < 10) {
        await new Promise(resolve => setTimeout(resolve, 100));
        container = this.$refs.pptxContainer;
        attempts++;
      }

      // 如果Vue的ref容器创建失败，使用临时容器方案
      if (!container) {
        const tempContainer = document.createElement('div');
        tempContainer.className = 'pptx-content';
        const pptxWrapper = document.querySelector('.document-viewer__content');
        if (pptxWrapper) {
          pptxWrapper.appendChild(tempContainer);
          container = tempContainer;
          console.log("✅ 使用临时PPTX容器");
        } else {
          throw new Error("无法创建PPTX预览容器");
        }
      }

      try {
        // 先清空容器
        container.innerHTML = '';
        
        // 创建OfficeViewer实例并渲染
        const officeViewer = new window.OfficeViewer();
        
        // 处理不同类型的文件输入
        let fileToRender = this.file;
        if (this.file instanceof ArrayBuffer) {
          // 如果是ArrayBuffer，需要转换为File对象
          const fileName = this.fileName || 'document.pptx';
          fileToRender = new File([this.file], fileName, { 
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' 
          });
          console.log('转换ArrayBuffer为File对象:', fileName);
        }
        
        await officeViewer.renderPPTXWithOfficeViewer(fileToRender, container);
        
        console.log("✅ PPTX预览渲染完成");
      } catch (error) {
        console.error("PPTX预览失败:", error);
        // 如果OfficeViewer失败，显示错误信息
        if (container) {
          container.innerHTML = `
            <div class="pptx-error">
              <div class="error-content">
                <h3>📊 PPTX预览</h3>
                <p>文件名: ${this.fileInfo.fileName}</p>
                <p>文件大小: ${this.fileInfo.fileSize}</p>
                <div class="error-message">
                  <p>⚠️ 预览加载失败: ${error.message}</p>
                  <p>这可能是由于文件格式复杂或损坏导致的</p>
                </div>
              </div>
            </div>
          `;
        }
        throw error;
      }
    },

    async renderExcel() {
      // 检查XLSX库是否已加载
      if (typeof window.XLSX === "undefined") {
        throw new Error("XLSX库未加载");
      }

      // 读取Excel文件
      const arrayBuffer =
        this.file instanceof File ? await this.file.arrayBuffer() : this.file;

      this.workbook = window.XLSX.read(arrayBuffer, { type: "array" });
      this.sheetNames = this.workbook.SheetNames;

      if (!this.sheetNames || this.sheetNames.length === 0) {
        throw new Error("Excel文件中没有找到工作表");
      }

      // 设置当前工作表为第一个
      this.currentSheetIndex = 0;
      this.updateCurrentSheetData();

      console.log("✅ Excel预览渲染完成");
    },

    updateCurrentSheetData() {
      if (!this.workbook || !this.sheetNames[this.currentSheetIndex]) return;

      const worksheet = this.workbook.Sheets[this.sheetNames[this.currentSheetIndex]];
      if (!worksheet) return;

      const range = window.XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      
      // 生成列标题（A, B, C, D...）
      const headers = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        headers.push(this.getColumnName(col));
      }

      // 提取所有数据行（包括原来的第一行）
      const rows = [];
      for (let row = range.s.r; row <= range.e.r; row++) {
        const rowData = [];
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = window.XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          const value = cell ? (cell.v || '') : '';
          rowData.push(String(value));
        }
        rows.push(rowData);
      }

      this.currentSheetData = { headers, rows };
    },

    // 获取Excel列名（A, B, C, D...）
    getColumnName(columnIndex) {
      let columnName = '';
      while (columnIndex >= 0) {
        columnName = String.fromCharCode(65 + (columnIndex % 26)) + columnName;
        columnIndex = Math.floor(columnIndex / 26) - 1;
      }
      return columnName;
    },

    switchSheet(index) {
      this.currentSheetIndex = index;
      this.updateCurrentSheetData();
    },

    exportCSV() {
      if (!this.workbook || !this.sheetNames[this.currentSheetIndex]) return;

      try {
        const currentSheet = this.workbook.Sheets[this.sheetNames[this.currentSheetIndex]];
        const csv = window.XLSX.utils.sheet_to_csv(currentSheet);
        
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${this.sheetNames[this.currentSheetIndex] || 'Sheet1'}.csv`;
        a.click();
        URL.revokeObjectURL(url);
        
        console.log('✅ CSV导出成功');
      } catch (error) {
        console.error('❌ CSV导出失败:', error);
        alert('导出CSV失败: ' + error.message);
      }
    },

    exportJSON() {
      if (!this.workbook || !this.sheetNames[this.currentSheetIndex]) return;

      try {
        const currentSheet = this.workbook.Sheets[this.sheetNames[this.currentSheetIndex]];
        
        // 构建按单元格位置为key的JSON格式
        const cellDataArray = [];
        
        // 遍历工作表中的所有单元格
        for (let cellAddress in currentSheet) {
          // 跳过以!开头的特殊属性（如 !ref, !margins 等）
          if (cellAddress[0] === '!') continue;
          
          const cell = currentSheet[cellAddress];
          // 获取单元格的值
          const cellValue = cell.v; // v 是单元格的原始值
          
          // 添加到数组中，格式为 { "A1": value }
          cellDataArray.push({
            [cellAddress]: cellValue
          });
        }
        
        const jsonStr = JSON.stringify(cellDataArray, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${this.sheetNames[this.currentSheetIndex] || 'Sheet1'}.json`;
        a.click();
        URL.revokeObjectURL(url);
        
        console.log('✅ JSON导出成功');
      } catch (error) {
        console.error('❌ JSON导出失败:', error);
        alert('导出JSON失败: ' + error.message);
      }
    },


    setupDownload() {
      window.downloadFile = () => {
        const url = URL.createObjectURL(this.file);
        const a = document.createElement("a");
        a.href = url;
        a.download = this.fileInfo.fileName;
        a.click();
        URL.revokeObjectURL(url);
      };
    },

    retry() {
      this.processFile();
    },

    getFileExtension(fileName) {
      return fileName.split(".").pop().toLowerCase();
    },

    formatFileSize(bytes) {
      if (bytes === 0) return "0 Bytes";
      const k = 1024;
      const sizes = ["Bytes", "KB", "MB", "GB"];
      const i = Math.floor(Math.log(bytes) / Math.log(k));
      return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
    },

    getFileIcon(type) {
      const icons = {
        pdf: "📄",
        docx: "📝",
        pptx: "📊",
        xlsx: "📈",
        xls: "📈",
        csv: "📋",
      };
      return icons[type] || "📄";
    },

    getFileTypeLabel(type) {
      const labels = {
        pdf: "PDF文档",
        docx: "Word文档",
        pptx: "PowerPoint演示文稿",
        xlsx: "Excel表格",
        xls: "Excel表格",
        csv: "CSV表格",
      };
      return labels[type] || "文档";
    },
  },
};
</script>

<style scoped>
.document-viewer {
  display: flex;
  flex-direction: column;
  height: 100%;
  background: #f5f5f5;
}

.document-viewer__header {
  background: white;
  border-bottom: 1px solid #e0e0e0;
  padding: 16px;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.document-viewer__header-info {
  display: flex;
  align-items: center;
  gap: 12px;
}

.file-icon {
  font-size: 24px;
}

.file-name {
  margin: 0 0 4px 0;
  font-size: 16px;
  font-weight: 600;
}

.file-meta {
  margin: 0;
  font-size: 14px;
  color: #666;
}

.file-meta span {
  margin-right: 12px;
}

.document-viewer__content {
  flex: 1;
  overflow: auto;
  padding: 20px;
}

.document-viewer__loading {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  height: 200px;
}

.loading-spinner {
  width: 40px;
  height: 40px;
  border: 4px solid #f3f3f3;
  border-top: 4px solid #3498db;
  border-radius: 50%;
  animation: spin 1s linear infinite;
  margin-bottom: 16px;
}

.loading-text {
  color: #666;
}

.document-viewer__error {
  text-align: center;
  padding: 40px;
}

.error-icon {
  font-size: 48px;
  margin-bottom: 16px;
}

.error-title {
  color: #e74c3c;
  margin-bottom: 8px;
}

.error-message {
  color: #666;
  margin-bottom: 20px;
}

.pdf-container {
  text-align: center;
}

.pdf-controls {
  margin-bottom: 20px;
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 16px;
}

.pdf-canvas {
  max-width: 100%;
  border: 1px solid #ddd;
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}

.document-info {
  text-align: center;
  padding: 40px;
  background: white;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.document-info h3 {
  margin-bottom: 20px;
  color: #333;
}

.document-info p {
  margin: 8px 0;
  color: #666;
}

.download-section {
  margin-top: 30px;
}

.btn {
  padding: 8px 16px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
}

.btn-primary {
  background: #007bff;
  color: white;
}

.btn-primary:hover {
  background: #0056b3;
}

.btn-secondary {
  background: #6c757d;
  color: white;
}

.btn-sm {
  padding: 6px 12px;
  font-size: 13px;
}

@keyframes spin {
  0% {
    transform: rotate(0deg);
  }
  100% {
    transform: rotate(360deg);
  }
}

/* DOCX预览样式 */
.docx-container {
  max-width: 800px;
  margin: 0 auto;
  background: white;
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
  border-radius: 8px;
  overflow: hidden;
}

.docx-content {
  padding: 40px;
  line-height: 1.6;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

.docx-content h1,
.docx-content h2,
.docx-content h3 {
  color: #333;
  margin-top: 24px;
  margin-bottom: 16px;
}

.docx-content p {
  margin-bottom: 12px;
  color: #444;
}

.docx-content img {
  max-width: 100%;
  height: auto;
  margin: 16px 0;
}

/* PPTX预览样式 */
.pptx-container {
  background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
  border-radius: 12px;
  padding: 0;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
  overflow: hidden;
}

.pptx-content {
  background: white !important;
  border-radius: 8px !important;
  margin: 16px !important;
  padding: 20px !important;
  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1) !important;
  min-height: 600px !important;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
}

/* 全局覆盖office-viewer的所有子元素 */
.pptx-content * {
  box-sizing: border-box !important;
}

/* 优化PPTX内部样式 */
.pptx-content .ppt-slide {
  background: white;
  border: 1px solid #e2e8f0;
  border-radius: 8px;
  padding: 20px;
  margin: 16px 0;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
  transition: all 0.3s ease;
}

.pptx-content .ppt-slide:hover {
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
  transform: translateY(-2px);
}

.pptx-content .ppt-slide h1,
.pptx-content .ppt-slide h2,
.pptx-content .ppt-slide h3 {
  color: #1e293b;
  font-weight: 600;
  margin-bottom: 12px;
  padding-bottom: 8px;
  border-bottom: 2px solid #3b82f6;
}

.pptx-content .ppt-slide h1 {
  font-size: 2rem;
  color: #1e40af;
}

.pptx-content .ppt-slide h2 {
  font-size: 1.5rem;
  color: #2563eb;
}

.pptx-content .ppt-slide h3 {
  font-size: 1.25rem;
  color: #3b82f6;
}

.pptx-content .ppt-slide p {
  color: #475569;
  line-height: 1.6;
  margin-bottom: 12px;
  font-size: 1rem;
}

.pptx-content .ppt-slide ul,
.pptx-content .ppt-slide ol {
  margin: 12px 0;
  padding-left: 24px;
}

.pptx-content .ppt-slide li {
  color: #475569;
  line-height: 1.6;
  margin-bottom: 6px;
  position: relative;
}

.pptx-content .ppt-slide ul li::before {
  content: "•";
  color: #3b82f6;
  font-weight: bold;
  position: absolute;
  left: -16px;
}

/* 优化页面导航样式 */
.pptx-content .ppt-navigation {
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  border-radius: 8px;
  padding: 12px 16px;
  margin: 16px 0;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 12px;
}

.pptx-content .ppt-navigation button {
  background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
  color: white;
  border: none;
  border-radius: 6px;
  padding: 8px 16px;
  font-size: 0.875rem;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s ease;
}

.pptx-content .ppt-navigation button:hover {
  background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
  transform: translateY(-1px);
  box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
}

.pptx-content .ppt-navigation button:disabled {
  background: #e2e8f0;
  color: #94a3b8;
  cursor: not-allowed;
  transform: none;
  box-shadow: none;
}

.pptx-content .ppt-navigation span {
  color: #475569;
  font-weight: 500;
  font-size: 0.875rem;
}

/* 优化office-viewer生成的内容样式 */
.pptx-content .office-viewer-content {
  background: white !important;
  border-radius: 12px !important;
  padding: 24px !important;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1) !important;
  min-height: 500px !important;
}

.pptx-content .loading-message {
  text-align: center !important;
  padding: 60px 20px !important;
  color: #64748b !important;
}

.pptx-content .loading-spinner {
  width: 40px !important;
  height: 40px !important;
  border: 4px solid #e2e8f0 !important;
  border-top: 4px solid #3b82f6 !important;
  border-radius: 50% !important;
  margin: 0 auto 16px !important;
  animation: spin 1s linear infinite !important;
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

.pptx-content .slide-container,
.pptx-content .ppt-slide-container {
  background: white !important;
  border: 2px solid #e2e8f0 !important;
  border-radius: 12px !important;
  margin: 20px 0 !important;
  padding: 24px !important;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08) !important;
  transition: all 0.3s ease !important;
}

.pptx-content .slide-container:hover,
.pptx-content .ppt-slide-container:hover {
  border-color: #3b82f6 !important;
  box-shadow: 0 8px 32px rgba(59, 130, 246, 0.15) !important;
  transform: translateY(-4px) !important;
}

.pptx-content .slide-title,
.pptx-content .ppt-title,
.pptx-content h1,
.pptx-content h2,
.pptx-content h3 {
  color: #1e40af !important;
  font-size: 24px !important;
  font-weight: 700 !important;
  margin-bottom: 16px !important;
  padding-bottom: 12px !important;
  border-bottom: 3px solid #3b82f6 !important;
  line-height: 1.3 !important;
}

.pptx-content .slide-content,
.pptx-content .ppt-text,
.pptx-content p,
.pptx-content div {
  color: #374151 !important;
  line-height: 1.8 !important;
  font-size: 16px !important;
  margin-bottom: 12px !important;
}

.pptx-content .slide-text,
.pptx-content .text-content {
  background: #f8fafc !important;
  border-radius: 8px !important;
  padding: 16px !important;
  margin: 12px 0 !important;
  border-left: 4px solid #3b82f6 !important;
}

.pptx-content ul,
.pptx-content ol {
  padding-left: 24px !important;
  margin: 16px 0 !important;
}

.pptx-content li {
  color: #475569 !important;
  line-height: 1.6 !important;
  margin-bottom: 8px !important;
  font-size: 15px !important;
}

.pptx-content ul li::marker {
  color: #3b82f6 !important;
  font-size: 18px !important;
}

/* 优化图片显示 */
.pptx-content img {
  max-width: 100% !important;
  height: auto !important;
  border-radius: 8px !important;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1) !important;
  margin: 16px 0 !important;
}

/* 优化表格样式 */
.pptx-content table {
  width: 100% !important;
  border-collapse: collapse !important;
  margin: 16px 0 !important;
  background: white !important;
  border-radius: 8px !important;
  overflow: hidden !important;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1) !important;
}

.pptx-content th,
.pptx-content td {
  padding: 12px 16px !important;
  text-align: left !important;
  border-bottom: 1px solid #e2e8f0 !important;
  color: #374151 !important;
}

.pptx-content th {
  background: #f1f5f9 !important;
  color: #1e40af !important;
  font-weight: 600 !important;
}

/* 如果没有内容，显示美化的提示 */
.pptx-content:empty::after {
  content: "📄 正在加载PPTX内容...";
  display: block;
  text-align: center;
  color: #64748b;
  font-size: 18px;
  padding: 60px;
  background: #f8fafc;
  border-radius: 12px;
  border: 2px dashed #cbd5e1;
}

/* 优化office-viewer的工具栏 */
.pptx-content .toolbar,
.pptx-content .ppt-toolbar,
.pptx-content .office-toolbar {
  background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%) !important;
  border: 1px solid #cbd5e1 !important;
  border-radius: 12px !important;
  padding: 12px 16px !important;
  margin: 0 0 20px 0 !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 12px !important;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06) !important;
}

.pptx-content .toolbar button,
.pptx-content .ppt-toolbar button,
.pptx-content .office-toolbar button,
.pptx-content button {
  background: linear-gradient(135deg, #ffffff 0%, #f1f5f9 100%) !important;
  border: 1px solid #cbd5e1 !important;
  border-radius: 8px !important;
  padding: 8px 12px !important;
  font-size: 13px !important;
  font-weight: 500 !important;
  color: #475569 !important;
  cursor: pointer !important;
  transition: all 0.2s ease !important;
  min-width: auto !important;
  height: auto !important;
  margin: 0 2px !important;
}

.pptx-content .toolbar button:hover,
.pptx-content .ppt-toolbar button:hover,
.pptx-content .office-toolbar button:hover,
.pptx-content button:hover {
  background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
  color: white !important;
  border-color: #2563eb !important;
  transform: translateY(-1px) !important;
  box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;
}

.pptx-content .toolbar select,
.pptx-content .ppt-toolbar select,
.pptx-content .office-toolbar select,
.pptx-content select {
  background: white !important;
  border: 1px solid #cbd5e1 !important;
  border-radius: 6px !important;
  padding: 6px 10px !important;
  font-size: 13px !important;
  color: #475569 !important;
  cursor: pointer !important;
  transition: all 0.2s ease !important;
}

.pptx-content .toolbar select:focus,
.pptx-content .ppt-toolbar select:focus,
.pptx-content .office-toolbar select:focus,
.pptx-content select:focus {
  outline: none !important;
  border-color: #3b82f6 !important;
  box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1) !important;
}

/* 特殊处理顶部工具栏区域 */
.pptx-content > div:first-child {
  background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%) !important;
  border-radius: 12px !important;
  padding: 12px !important;
  margin-bottom: 16px !important;
  border: 1px solid #cbd5e1 !important;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06) !important;
}

.pptx-content > div:first-child * {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
}

/* 更强制的工具栏样式覆盖 */
.pptx-content input[type="button"],
.pptx-content input[type="submit"],
.pptx-content input[type="reset"] {
  background: linear-gradient(135deg, #ffffff 0%, #f1f5f9 100%) !important;
  border: 1px solid #cbd5e1 !important;
  border-radius: 8px !important;
  padding: 6px 12px !important;
  font-size: 13px !important;
  font-weight: 500 !important;
  color: #475569 !important;
  cursor: pointer !important;
  transition: all 0.2s ease !important;
  margin: 0 3px !important;
}

.pptx-content input[type="button"]:hover,
.pptx-content input[type="submit"]:hover,
.pptx-content input[type="reset"]:hover {
  background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
  color: white !important;
  border-color: #2563eb !important;
  transform: translateY(-1px) !important;
  box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;
}

/* 处理可能的表格布局工具栏 */
.pptx-content table[style*="width"] {
  width: 100% !important;
  border-collapse: separate !important;
  border-spacing: 0 !important;
  background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%) !important;
  border: 1px solid #cbd5e1 !important;
  border-radius: 12px !important;
  padding: 8px !important;
  margin-bottom: 16px !important;
}

.pptx-content table[style*="width"] td {
  padding: 4px 8px !important;
  border: none !important;
  text-align: center !important;
}

/* 处理可能的span或div形式的按钮 */
.pptx-content span[onclick],
.pptx-content div[onclick] {
  background: linear-gradient(135deg, #ffffff 0%, #f1f5f9 100%) !important;
  border: 1px solid #cbd5e1 !important;
  border-radius: 8px !important;
  padding: 6px 12px !important;
  font-size: 13px !important;
  font-weight: 500 !important;
  color: #475569 !important;
  cursor: pointer !important;
  transition: all 0.2s ease !important;
  display: inline-block !important;
  margin: 0 3px !important;
  user-select: none !important;
}

.pptx-content span[onclick]:hover,
.pptx-content div[onclick]:hover {
  background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
  color: white !important;
  border-color: #2563eb !important;
  transform: translateY(-1px) !important;
  box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;
}

/* 强制覆盖所有可能的内联样式 */
.pptx-content [style] {
  border-radius: 8px !important;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1) !important;
}

.pptx-content [style*="font-size"] {
  font-size: 16px !important;
  line-height: 1.6 !important;
}

.pptx-content [style*="color"] {
  color: #374151 !important;
}

.pptx-content [style*="background"] {
  background: white !important;
  border: 1px solid #e2e8f0 !important;
  border-radius: 8px !important;
  padding: 16px !important;
  margin: 12px 0 !important;
}

/* 特殊处理文本内容 */
.pptx-content span,
.pptx-content div:not(.office-viewer-content):not(.loading-message) {
  color: #374151 !important;
  font-size: 16px !important;
  line-height: 1.6 !important;
  padding: 4px 0 !important;
}

/* 处理可能的标题元素 */
.pptx-content [style*="font-weight: bold"],
.pptx-content [style*="font-weight:bold"],
.pptx-content b,
.pptx-content strong {
  color: #1e40af !important;
  font-size: 20px !important;
  font-weight: 700 !important;
  display: block !important;
  margin: 16px 0 12px 0 !important;
  padding-bottom: 8px !important;
  border-bottom: 2px solid #3b82f6 !important;
}

/* Excel预览样式 - 使用原项目的完整样式，这里只保留必要的覆盖 */

.excel-header {
  padding: 16px 20px;
  background: #f8f9fa;
  border-bottom: 1px solid #e0e0e0;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.excel-header h3 {
  margin: 0;
  color: #333;
}

.sheet-selector {
  display: flex;
  align-items: center;
  gap: 8px;
}

.sheet-selector label {
  font-size: 14px;
  color: #666;
}

.sheet-selector select {
  padding: 4px 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 14px;
}

.excel-content {
  padding: 20px;
  overflow: auto;
  max-height: 600px;
}

.excel-content table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
}

.excel-content th,
.excel-content td {
  border: 1px solid #ddd;
  padding: 8px 12px;
  text-align: left;
}

.excel-content th {
  background: #f5f5f5;
  font-weight: 600;
  color: #333;
}

.excel-content tr:nth-child(even) {
  background: #f9f9f9;
}

.excel-content tr:hover {
  background: #f0f8ff;
}

/* 错误样式 */
.docx-error,
.pptx-error,
.excel-error {
  text-align: center;
  padding: 40px;
  background: white;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.error-content h3 {
  margin-bottom: 20px;
  color: #333;
}

.error-content p {
  margin: 8px 0;
  color: #666;
}

.error-message {
  margin-top: 20px;
  padding: 16px;
  background: #fff3cd;
  border: 1px solid #ffeaa7;
  border-radius: 4px;
  color: #856404;
}
</style>