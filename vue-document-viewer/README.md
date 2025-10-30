# Vue Document Viewer

一个基于Vue 2的文档预览组件库，支持在线预览PDF、DOCX、PPTX、XLSX等多种格式的文档，无需服务器依赖。

## ✨ 特性

- 🎯 **多格式支持**: PDF、DOCX、PPTX、XLSX、CSV
- 🚀 **纯前端实现**: 无需服务器，完全在浏览器中处理
- 📱 **响应式设计**: 适配桌面和移动设备
- 🎨 **现代化UI**: 美观的界面设计和交互体验
- 🔧 **易于集成**: 支持Vue组件和全局安装
- 📦 **轻量级**: 优化的打包体积
- 🌐 **多种使用方式**: 本地文件上传、URL预览、直接预览

## 📦 安装

```bash
npm install @anohanafes/vue-document-viewer
# 或
yarn add @anohanafes/vue-document-viewer
```

## 🚀 快速开始

### 全局注册

```javascript
import Vue from 'vue'
import VueDocumentViewer from '@anohanafes/vue-document-viewer'

Vue.use(VueDocumentViewer)
```

### 按需引入

```javascript
import { DocumentViewer, FileUploader } from '@anohanafes/vue-document-viewer'

export default {
  components: {
    DocumentViewer,
    FileUploader
  }
}
```

### 基础用法

```vue
<template>
  <div>
    <!-- 文件上传器 -->
    <FileUploader 
      @file-selected="handleFileSelected"
      :max-size="52428800"
    />
    
    <!-- 文档预览器 -->
    <DocumentViewer 
      :file="selectedFile"
      :file-name="fileName"
    />
  </div>
</template>

<script>
export default {
  data() {
    return {
      selectedFile: null,
      fileName: ''
    }
  },
  methods: {
    handleFileSelected(file) {
      this.selectedFile = file
      this.fileName = file.name
    }
  }
}
</script>
```

## 📚 组件文档

### DocumentViewer

文档预览核心组件

#### Props

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| file | File \| ArrayBuffer | null | 要预览的文件对象 |
| fileName | String | '' | 文件名（当file为ArrayBuffer时必需） |
| hideHeader | Boolean | false | 是否隐藏头部信息 |

#### Events

| 事件名 | 说明 | 回调参数 |
|--------|------|----------|
| close | 关闭预览时触发 | - |

### FileUploader

文件上传组件

#### Props

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| maxSize | Number | 52428800 | 最大文件大小（字节） |
| acceptedTypes | Array | ['pdf', 'docx', 'pptx', 'xlsx', 'csv'] | 支持的文件类型 |
| autoPreview | Boolean | false | 是否自动预览 |

#### Events

| 事件名 | 说明 | 回调参数 |
|--------|------|----------|
| file-selected | 文件选择时触发 | file: File |
| file-cleared | 文件清除时触发 | - |
| preview-file | 预览文件时触发 | file: File |

### UrlInput

URL输入组件

#### Props

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| placeholder | String | '请输入文档URL' | 输入框占位符 |

#### Events

| 事件名 | 说明 | 回调参数 |
|--------|------|----------|
| url-submit | URL提交时触发 | url: String |

## 🎨 主题定制

组件使用CSS变量，可以轻松定制主题：

```css
:root {
  --primary-color: #2563eb;
  --success-color: #16a34a;
  --warning-color: #ea580c;
  --error-color: #dc2626;
  --border-radius: 8px;
  --box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}
```

## 🔧 高级用法

### 自定义样式

```vue
<template>
  <DocumentViewer 
    :file="file"
    class="custom-viewer"
  />
</template>

<style>
.custom-viewer {
  border: 2px solid #e5e7eb;
  border-radius: 12px;
  overflow: hidden;
}
</style>
```

### 处理大文件

```javascript
// 对于大文件，建议添加加载提示
methods: {
  async handleLargeFile(file) {
    if (file.size > 10 * 1024 * 1024) { // 10MB
      this.showLoading = true
    }
    
    try {
      await this.processFile(file)
    } finally {
      this.showLoading = false
    }
  }
}
```

## 📱 移动端支持

组件完全支持移动端，包括：
- 触摸手势支持
- 响应式布局
- 移动端优化的UI

## 🌐 浏览器兼容性

- Chrome >= 60
- Firefox >= 60
- Safari >= 12
- Edge >= 79

## 📄 支持的文件格式

| 格式 | 支持功能 |
|------|----------|
| PDF | ✅ 在线预览、翻页、缩放 |
| DOCX | ✅ 在线预览、样式保持 |
| PPTX | ✅ 幻灯片预览、翻页 |
| XLSX | ✅ 表格预览、多工作表、导出 |
| CSV | ✅ 表格预览、导出 |

## 🛠️ 开发

```bash
# 克隆项目
git clone https://github.com/anohanafes/offline-document-viewer.git
cd offline-document-viewer/vue-document-viewer

# 安装依赖
npm install

# 开发模式
npm run serve

# 构建库
npm run build:lib

# 构建演示
npm run build:demo
```

## 📝 更新日志

### 1.0.0
- 🎉 首次发布
- ✅ 支持PDF、DOCX、PPTX、XLSX预览
- ✅ 现代化UI设计
- ✅ 完整的TypeScript支持

## 🤝 贡献

欢迎提交Issue和Pull Request！

## 📄 许可证

[MIT License](./LICENSE)

## 🙏 致谢

- [PDF.js](https://mozilla.github.io/pdf.js/) - PDF渲染
- [Mammoth.js](https://github.com/mwilliamson/mammoth.js/) - DOCX转换
- [SheetJS](https://sheetjs.com/) - Excel处理
- [JSZip](https://stuk.github.io/jszip/) - ZIP文件处理