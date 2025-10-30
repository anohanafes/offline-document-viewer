# Changelog

All notable changes to this project will be documented in this file.

## [2.0.0] - 2024-12-19

### 🎉 Major Release - Complete Rewrite

### ✨ Added
- **Multi-Mode Preview System**
  - `index.html` - Local file upload and preview
  - `url-viewer.html` - URL-based document preview
  - `direct-viewer.html` - Embed-friendly direct preview
  
- **Enhanced Format Support**
  - Excel files (`.xlsx`, `.xls`) with multi-worksheet support
  - CSV files with UTF-8 encoding and smart detection
  - Improved PPTX rendering with custom engine
  
- **Smart Resource Management**
  - Dynamic on-demand loading system
  - Intelligent caching and preloading
  - Resource usage optimization (up to 80% bandwidth savings)
  
- **Custom Rendering Engines**
  - `office-viewer.js` - Self-developed PPTX renderer
  - `excel-viewer.js` - Advanced Excel preview component
  - `resource-manager.js` - Smart resource loader
  - `document-processor-enhancer.js` - Processing enhancer

### 🔧 Improved
- **Performance Optimization**
  - Reduced initial loading time by 95%
  - On-demand loading for all document types
  - Smart preloading based on usage patterns
  
- **User Experience**
  - Clean, modern interface design
  - Mobile-responsive layouts
  - Keyboard navigation support
  - Error handling and recovery

### 🔄 Changed
- **Architecture Redesign**
  - Modular component structure
  - Event-driven processing system
  - Separation of concerns
  
- **No More Build Process**
  - Removed npm build requirements
  - All dependencies included
  - Zero-configuration deployment

### 🗑️ Removed
- Build scripts and npm dependencies
- Drag-and-drop upload (simplified to click-to-select)
- Outdated pptx.js library

---

## [1.0.0] - Initial Release

### ✨ Features
- Basic PDF document viewing
- DOCX to HTML conversion
- Simple PPTX information display
- Local file upload interface
- Responsive design foundation

### 🛠️ Tech Stack
- PDF.js for PDF rendering
- Mammoth.js for DOCX conversion
- JSZip for archive handling
- jQuery for DOM manipulation
