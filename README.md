# 📄 离线文档预览器

<div align="center">

一个功能完整的多格式文档在线预览解决方案，支持完全离线运行和URL在线预览

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![GitHub stars](https://img.shields.io/github/stars/anohanafes/offline-document-viewer?style=social)](https://github.com/anohanafes/offline-document-viewer/stargazers)
[![npm version](https://img.shields.io/npm/v/@anohanafes/offline-document-viewer.svg)](https://www.npmjs.com/package/@anohanafes/offline-document-viewer)

[🔗 在线演示](https://anohanafes.github.io/offline-document-viewer/) | [📖 文档](https://github.com/anohanafes/offline-document-viewer/wiki) | [🐛 报告问题](https://github.com/anohanafes/offline-document-viewer/issues)

</div>

## ✨ 特性

- 🔒 **双重模式** - 支持完全离线运行和URL在线预览
- 📄 **全格式支持** - PDF、DOCX、PPTX、XLSX、CSV完整预览
- ⚡ **按需加载** - 智能动态加载，只加载必需资源
- 🎯 **多种预览方式** - 本地文件上传、URL链接、直接嵌入预览
- 📱 **响应式设计** - 完美支持桌面和移动设备
- ⌨️ **键盘快捷键** - 丰富的快捷键操作
- 🚀 **开箱即用** - 无需构建，直接运行

## 🚀 快速开始

### 方式一：直接使用（推荐）

1. **克隆项目**
```bash
git clone https://github.com/anohanafes/offline-document-viewer.git
cd offline-document-viewer
```

2. **启动服务**
```bash
# Python 3 (推荐)
python -m http.server 8080

# Node.js 
npx http-server -p 8080 -c-1

# PHP
php -S localhost:8080

# 或使用任何静态文件服务器
```

3. **访问应用**
```
http://localhost:8080/index.html          # 本地文件预览
http://localhost:8080/url-viewer.html     # URL文档预览  
http://localhost:8080/direct-viewer.html  # 直接预览（嵌入式）
```

### 方式二：直接打开
双击 `index.html` 文件即可在浏览器中直接使用（部分功能可能受限）

## 📋 支持的文件格式

| 格式 | 扩展名 | 功能特性 | 状态 |
|------|--------|----------|------|
| **PDF** | `.pdf` | 完整预览、缩放、翻页、搜索 | ✅ 完全支持 |
| **Word** | `.docx`, `.doc` | HTML转换预览、样式保留、缩放 | ✅ 完全支持 |
| **PowerPoint** | `.pptx`, `.ppt` | 幻灯片预览、样式渲染、导航 | ✅ 完全支持 |
| **Excel** | `.xlsx`, `.xls` | 多工作表、数据预览、JSON导出 | ✅ 完全支持 |
| **CSV** | `.csv` | UTF-8支持、表格显示、数据导出 | ✅ 完全支持 |

## 🛠️ 技术架构

### 核心技术栈
- **PDF.js** (Mozilla) - PDF文档渲染引擎
- **Mammoth.js** - DOCX到HTML转换
- **JSZip** - Office文档压缩包解析
- **SheetJS (XLSX)** - Excel文件处理
- **jQuery** - DOM操作和事件处理

### 自研组件
- **Office-Viewer** - 自研PPTX渲染引擎，支持样式和布局
- **Excel-Viewer** - 自研Excel预览组件，支持多工作表
- **Dynamic Resource Manager** - 智能按需资源加载系统
- **Document Processor Enhancer** - 文档处理增强器

### 项目结构
```
offline-document-viewer/
├── index.html                    # 🏠 本地文件预览主页
├── url-viewer.html               # 🔗 URL文档预览页面  
├── direct-viewer.html            # 📺 直接预览页面（嵌入式）
├── package.json                  # 📦 项目配置文件
├── README.md                     # 📚 项目说明文档
├── LICENSE                       # ⚖️ 开源许可证
├── .gitignore                    # 🚫 Git忽略文件
├── css/                          # 🎨 样式文件目录
│   ├── styles.css               #     主样式文件
│   ├── pptx-styles.css          #     PPTX专用样式
│   └── excel-styles.css         #     Excel专用样式
├── js/                          # ⚙️ JavaScript逻辑
│   ├── document-viewer.js       #     主应用逻辑
│   ├── resource-manager.js      #     动态资源管理器
│   └── document-processor-enhancer.js  # 文档处理增强器
└── lib/                         # 📚 第三方库文件（已包含）
    ├── jquery.min.js            #     jQuery库
    ├── jszip.min.js             #     ZIP文件处理
    ├── pdf.min.js               #     PDF渲染引擎
    ├── pdf.worker.js            #     PDF处理Worker
    ├── mammoth.min.js           #     DOCX转换器
    ├── xlsx.full.min.js         #     Excel处理库
    ├── office-viewer.js         #     自研PPTX渲染器
    └── excel-viewer.js          #     自研Excel预览器
```

## 🎮 使用说明

### 📁 本地文件预览 (`index.html`)
1. **点击选择文件** - 打开文件选择器
2. **选择文档** - 支持PDF、DOCX、PPTX、XLSX、CSV
3. **自动预览** - 文档将自动加载并显示

### 🔗 URL文档预览 (`url-viewer.html`)
1. **输入文档URL** - 在输入框中粘贴文档链接
2. **点击加载预览** - 自动下载并渲染文档
3. **跨域支持** - 需要目标服务器支持CORS

**示例URL格式：**
```
url-viewer.html?url=https://example.com/document.pdf
```

### 📺 直接预览模式 (`direct-viewer.html`)
**无UI干扰的纯预览模式，适合嵌入到其他系统**

```html
<!-- 嵌入到您的网页中 -->
<iframe src="direct-viewer.html?url=YOUR_DOCUMENT_URL" 
        width="100%" height="600px">
</iframe>
```

**特点：**
- ✅ 无文件信息面板
- ✅ 浮动控制栏  
- ✅ 全屏沉浸式体验
- ✅ 适合系统集成

### ⌨️ 键盘快捷键
- `←/→` - PDF翻页 / PPTX切换幻灯片
- `+/-` - 放大/缩小
- `ESC` - 退出全屏模式
- `Space` - PDF下一页

## 🔧 自定义配置

### 🎨 界面主题定制
编辑 `css/styles.css` 自定义外观：

```css
:root {
    --primary-color: #2196F3;
    --secondary-color: #1976D2;
    --background: #f5f5f5;
}
```

### 📋 按需加载控制
在浏览器控制台中控制预加载行为：

```javascript
// 禁用智能预加载，完全按需加载
localStorage.setItem('documentViewer_disablePreload', 'true');

// 启用智能预加载
localStorage.setItem('documentViewer_disablePreload', 'false');
```

## 📱 设备兼容性

### 浏览器支持
- **推荐**: Chrome 90+, Firefox 88+, Safari 14+, Edge 90+
- **移动端**: iOS Safari 14+, Chrome Mobile 90+
- **必需特性**: ES6, Canvas API, File API, Fetch API

### 性能特点
- **PDF**: 渲染速度最快，大文件支持好
- **DOCX**: 转换速度快，样式保留度高
- **PPTX**: 自研引擎，支持复杂样式和布局
- **Excel**: 支持大型表格，多工作表预览
- **按需加载**: 只加载当前文档类型需要的库

### 文件大小建议
- **推荐**: < 50MB 获得最佳体验
- **支持**: 理论上无限制，取决于浏览器内存
- **优化**: 启用按需加载，减少初始加载时间

## 🔒 安全特性

- ✅ **本地处理优先** - 本地文件完全在浏览器中处理
- ✅ **URL模式透明** - URL预览时文件临时下载，不存储
- ✅ **浏览器沙盒** - 运行在浏览器安全沙盒环境中  
- ✅ **无数据收集** - 不收集、不存储任何用户数据
- ✅ **开源透明** - 所有代码公开，安全可审计

## 📦 安装

### NPM 安装

```bash
npm install @anohanafes/offline-document-viewer
```

### 直接使用

```bash
# 克隆项目
git clone https://github.com/anohanafes/offline-document-viewer.git
cd offline-document-viewer

# 启动服务
python -m http.server 8080
# 或
npx http-server -p 8080 -c-1
```

## 🤝 贡献指南

欢迎各种形式的贡献！🎉

### 📝 如何贡献
1. **Fork** 本项目
2. **创建** 特性分支 (`git checkout -b feature/AmazingFeature`)
3. **提交** 更改 (`git commit -m 'Add some AmazingFeature'`)
4. **推送** 到分支 (`git push origin feature/AmazingFeature`)
5. **创建** Pull Request

### 🐛 报告问题
- 使用 [GitHub Issues](https://github.com/anohanafes/offline-document-viewer/issues)
- 提供详细的重现步骤
- 包含浏览器版本和操作系统信息

### 💡 功能建议
- 提交 [Feature Request](https://github.com/anohanafes/offline-document-viewer/issues/new)
- 描述期望的功能和使用场景
- 解释为什么这个功能很有用

## 🌟 项目亮点

### 🚀 性能优化
- **按需加载** - 只加载当前需要的库，节省带宽和加载时间
- **智能缓存** - 资源加载后缓存，切换文档类型更快
- **渐进增强** - 基础功能优先，高级功能按需启用

### 🔧 开发友好
- **零配置** - 开箱即用，无需构建工具
- **模块化** - 清晰的代码结构，易于扩展
- **事件驱动** - 灵活的事件系统，支持自定义

### 🎯 用户体验
- **多模式支持** - 本地上传、URL预览、直接嵌入
- **响应式设计** - 完美适配各种设备屏幕
- **无障碍访问** - 支持键盘导航和屏幕阅读器

## 🙏 致谢

感谢以下优秀的开源项目：
- [PDF.js](https://github.com/mozilla/pdf.js) - Mozilla PDF渲染引擎
- [Mammoth.js](https://github.com/mwilliamson/mammoth.js) - DOCX转HTML转换器
- [JSZip](https://github.com/Stuk/jszip) - JavaScript ZIP处理库
- [SheetJS](https://github.com/SheetJS/sheetjs) - Excel文件处理库
- [jQuery](https://github.com/jquery/jquery) - DOM操作库

## 📄 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件

## 📞 联系我们

- 📧 **邮件联系**: 519855937@qq.com

---

<div align="center">

**⭐ 如果这个项目对您有帮助，请给个Star支持一下！**

Made with ❤️ by developers, for developers

</div>
