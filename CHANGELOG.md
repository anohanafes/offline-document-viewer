# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/lang/zh-CN/).

---

## [1.0.0] - 2024-10-31

### 🎉 首次发布

一个功能完整的多格式文档在线预览解决方案，支持完全离线运行和URL在线预览。

### ✨ 核心功能

#### 📂 多模式预览系统
- **本地文件预览** (`index.html`) - 支持本地文件上传和预览
- **URL文档预览** (`url-viewer.html`) - 支持通过URL链接预览远程文档
- **直接预览模式** (`direct-viewer.html`) - 无UI干扰的嵌入式预览，适合系统集成

#### 📄 全格式支持
- **PDF** - 完整预览、缩放、翻页、搜索功能
- **Word** (`.docx`, `.doc`) - HTML转换预览、样式保留、缩放
- **PowerPoint** (`.pptx`, `.ppt`) - 幻灯片预览、样式渲染、导航
- **Excel** (`.xlsx`, `.xls`) - 多工作表支持、数据预览、JSON导出
- **CSV** - UTF-8支持、表格显示、数据导出

#### ⚡ 智能资源管理
- **按需加载** - 动态加载，只加载当前文档类型需要的库
- **智能缓存** - 资源加载后缓存，切换文档类型更快
- **预加载优化** - 基于使用模式的智能预加载
- **性能优化** - 相比全量加载节省高达 80% 的带宽

#### 🎨 用户体验
- **响应式设计** - 完美支持桌面和移动设备
- **键盘快捷键** - 丰富的快捷键操作支持
- **现代化界面** - 简洁美观的用户界面
- **无障碍访问** - 支持键盘导航和屏幕阅读器

### 🛠️ 技术架构

#### 核心技术栈
- **PDF.js** (Mozilla) - PDF文档渲染引擎
- **Mammoth.js** - DOCX到HTML转换
- **JSZip** - Office文档压缩包解析
- **SheetJS (XLSX)** - Excel文件处理
- **jQuery** - DOM操作和事件处理

#### 自研组件
- **Office-Viewer** - 自研PPTX渲染引擎，支持复杂样式和布局
- **Excel-Viewer** - 自研Excel预览组件，支持多工作表切换
- **Resource Manager** - 智能按需资源加载系统
- **Document Processor Enhancer** - 文档处理增强器

### 🔒 安全特性
- ✅ 本地文件完全在浏览器中处理，不上传服务器
- ✅ URL预览时文件临时下载，不存储
- ✅ 运行在浏览器安全沙盒环境中
- ✅ 不收集、不存储任何用户数据
- ✅ 开源透明，所有代码可审计

### 📦 部署方式
- **零配置运行** - 无需构建工具，开箱即用
- **静态文件部署** - 可部署到任何静态文件服务器
- **双击运行** - 支持直接在浏览器中打开使用

### 🌐 浏览器兼容性
- Chrome 90+
- Firefox 88+
- Safari 14+
- Edge 90+
- 移动端浏览器支持

---

## 未来计划

- [ ] 添加更多文档格式支持（如 Markdown、TXT）
- [ ] 增强打印功能
- [ ] 添加文档水印功能
- [ ] 支持文档批量预览
- [ ] 性能优化和体验改进

---

<div align="center">

**感谢使用离线文档预览器！**

如有问题或建议，欢迎提交 Issue：
- GitHub: https://github.com/anohanafes/offline-document-viewer/issues
- Gitee: https://gitee.com/wang-qiuning/offline-document-viewer/issues

</div>
