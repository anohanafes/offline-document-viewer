/**
 * =====================================================
 * 离线文档预览器 - 重构版本
 * 支持PDF、DOCX、PPTX、Excel格式的完全离线预览
 * =====================================================
 */

// ====================================================================================
// 工具类：负责通用工具函数
// ====================================================================================
class DocumentUtils {
    /**
     * 格式化文件大小显示
     * @param {number} bytes - 文件大小（字节）
     * @returns {string} 格式化后的文件大小字符串
     */
    static formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    /**
     * 获取文件扩展名
     * @param {string} fileName - 文件名
     * @returns {string} 小写的文件扩展名
     */
    static getFileExtension(fileName) {
        return fileName.split('.').pop().toLowerCase();
    }

    /**
     * 检查是否为支持的文件格式
     * @param {string} extension - 文件扩展名
     * @returns {boolean} 是否支持该格式
     */
    static isSupportedFormat(extension) {
        const supportedFormats = ['pdf', 'docx', 'pptx', 'xlsx', 'xls', 'xlsm', 'xlsb', 'csv'];
        return supportedFormats.includes(extension);
    }

    /**
     * 动态加载脚本文件
     * @param {string} src - 脚本文件路径
     * @returns {Promise} 加载完成的Promise
     */
    static loadScript(src) {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    /**
     * 动态加载CSS文件
     * @param {string} href - CSS文件路径
     * @returns {Promise} 加载完成的Promise
     */
    static loadCSS(href) {
        return new Promise((resolve, reject) => {
            const link = document.createElement('link');
            link.rel = 'stylesheet';
            link.href = href;
            link.onload = resolve;
            link.onerror = reject;
            document.head.appendChild(link);
        });
    }
}

// ====================================================================================
// UI控制器：负责界面显示和交互
// ====================================================================================
class UIController {
    constructor() {
        this.currentZoom = 1.0;
    }

    /**
     * 显示文件信息
     * @param {File} file - 选中的文件对象
     */
    showFileInfo(file) {
        const fileInfo = document.getElementById('documentInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');
        const fileType = document.getElementById('fileType');
        
        fileName.textContent = file.name;
        fileSize.textContent = DocumentUtils.formatFileSize(file.size);
        fileType.textContent = file.type || 'Unknown';
        
        fileInfo.style.display = 'flex';
    }
    
    /**
     * 显示加载动画
     */
    showLoading() {
        document.getElementById('loadingOverlay').style.display = 'flex';
        document.getElementById('viewerControls').style.display = 'none';
    }
    
    /**
     * 隐藏加载动画
     */
    hideLoading() {
        document.getElementById('loadingOverlay').style.display = 'none';
    }
    
    /**
     * 显示错误信息
     * @param {string} message - 错误消息
     */
    showError(message) {
        this.hideLoading();
        const container = document.getElementById('viewerContainer');
        container.innerHTML = `
            <div class="error-message">
                <h3>❌ 处理失败</h3>
                <p>${message}</p>
                <button class="btn" onclick="location.reload()">重新开始</button>
            </div>
        `;
    }
    
    /**
     * 显示成功消息（3秒后自动消失）
     * @param {string} message - 成功消息
     */
    showSuccess(message) {
        const container = document.getElementById('viewerContainer');
        const successDiv = document.createElement('div');
        successDiv.className = 'success-message';
        successDiv.innerHTML = `<p>✅ ${message}</p>`;
        
        // 3秒后自动消失
        setTimeout(() => {
            if (successDiv.parentNode) {
                successDiv.remove();
            }
        }, 3000);
    }
    
    /**
     * 显示警告信息（5秒后自动消失）
     * @param {Array} messages - 警告消息数组
     */
    showWarnings(messages) {
        const warningDiv = document.createElement('div');
        warningDiv.className = 'warning-message';
        warningDiv.style.cssText = `
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            color: #856404;
            padding: 15px;
            margin: 10px 20px;
            border-radius: 6px;
            font-size: 14px;
        `;
        
        let warningText = '⚠️ 转换过程中的注意事项：<br>';
        messages.forEach(msg => {
            warningText += `• ${msg.message}<br>`;
        });
        
        warningDiv.innerHTML = warningText;
        document.getElementById('viewerContainer').prepend(warningDiv);
        
        // 5秒后自动消失
        setTimeout(() => {
            if (warningDiv.parentNode) {
                warningDiv.remove();
            }
        }, 5000);
    }

    /**
     * 隐藏所有控制面板
     */
    hideAllControls() {
        const controls = ['pdfControls', 'docxControls', 'pptxControls'];
        controls.forEach(controlId => {
            const control = document.getElementById(controlId);
            if (control) control.style.display = 'none';
        });
    }

    /**
     * 显示指定控制面板
     * @param {string} controlType - 控制面板类型 ('pdf', 'docx', 'pptx', 'excel')
     */
    showControls(controlType) {
        const controls = document.getElementById('viewerControls');
        
        // 隐藏所有控制面板
        this.hideAllControls();
        
        if (controlType === 'excel') {
            // Excel控制功能已集成在ExcelViewer内部，隐藏控制面板
            controls.style.display = 'none';
            return;
        }
        
        // 显示对应的控制面板
        controls.style.display = 'block';
        const targetControl = document.getElementById(`${controlType}Controls`);
        if (targetControl) {
            targetControl.style.display = 'flex';
        }
    }
}

// ====================================================================================
// 文档渲染器：包含PDF、DOCX、PPTX、Excel等所有渲染器
// ====================================================================================
class PDFRenderer {
    constructor(uiController) {
        this.ui = uiController;
        this.pdfDoc = null;
        this.currentPage = 1;
        this.totalPages = 1;
        this.currentZoom = 1.0;
    }

    async render(file) {
        try {
            // console.log('📄 开始渲染PDF文档');
            const arrayBuffer = await file.arrayBuffer();
            this.pdfDoc = await pdfjsLib.getDocument(arrayBuffer).promise;
            this.totalPages = this.pdfDoc.numPages;
            this.currentPage = 1;
            
            this._setupViewer();
            await this._renderPage(this.currentPage);
            this._setupControls();
            this.ui.hideLoading();
            
            // console.log('✅ PDF文档渲染完成');
        } catch (error) {
            throw new Error(`PDF加载失败: ${error.message}`);
        }
    }
    
    _setupViewer() {
        const container = document.getElementById('viewerContainer');
        container.innerHTML = `
            <div class="pdf-container">
                <canvas id="pdfCanvas" class="pdf-canvas"></canvas>
                <div class="progress-bar">
                    <div class="progress-fill" id="pdfProgress" style="width: ${(this.currentPage / this.totalPages) * 100}%"></div>
                </div>
            </div>
        `;
    }
    
    async _renderPage(pageNum) {
        const page = await this.pdfDoc.getPage(pageNum);
        const canvas = document.getElementById('pdfCanvas');
        const context = canvas.getContext('2d');
        
        const viewport = page.getViewport({ scale: this.currentZoom });
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        
        await page.render({
            canvasContext: context,
            viewport: viewport
        }).promise;
        
        this._updateProgress(pageNum);
    }

    _updateProgress(pageNum) {
        const progress = document.getElementById('pdfProgress');
        if (progress) {
            progress.style.width = `${(pageNum / this.totalPages) * 100}%`;
        }
    }
    
    _setupControls() {
        this.ui.showControls('pdf');
        this._updatePageInfo();
        
        document.getElementById('prevPage').onclick = () => this.previousPage();
        document.getElementById('nextPage').onclick = () => this.nextPage();
        document.getElementById('zoomOut').onclick = () => this.zoomOut();
        document.getElementById('zoomIn').onclick = () => this.zoomIn();
        document.getElementById('fitWidth').onclick = () => this.fitWidth();
    }
    
    _updatePageInfo() {
        document.getElementById('pageInfo').textContent = `第 ${this.currentPage} 页，共 ${this.totalPages} 页`;
        document.getElementById('zoomLevel').textContent = `${Math.round(this.currentZoom * 100)}%`;
        document.getElementById('prevPage').disabled = this.currentPage <= 1;
        document.getElementById('nextPage').disabled = this.currentPage >= this.totalPages;
    }
    
    async previousPage() {
        if (this.currentPage > 1) {
            this.currentPage--;
            await this._renderPage(this.currentPage);
            this._updatePageInfo();
        }
    }
    
    async nextPage() {
        if (this.currentPage < this.totalPages) {
            this.currentPage++;
            await this._renderPage(this.currentPage);
            this._updatePageInfo();
        }
    }
    
    async zoomIn() {
        this.currentZoom = Math.min(this.currentZoom * 1.2, 3.0);
        await this._renderPage(this.currentPage);
        this._updatePageInfo();
    }
    
    async zoomOut() {
        this.currentZoom = Math.max(this.currentZoom / 1.2, 0.5);
        await this._renderPage(this.currentPage);
        this._updatePageInfo();
    }
    
    async fitWidth() {
        const canvas = document.getElementById('pdfCanvas');
        const container = canvas.parentElement;
        const containerWidth = container.clientWidth - 40;
        
        const page = await this.pdfDoc.getPage(this.currentPage);
        const viewport = page.getViewport({ scale: 1.0 });
        
        this.currentZoom = containerWidth / viewport.width;
        await this._renderPage(this.currentPage);
        this._updatePageInfo();
    }
}

class DOCXRenderer {
    constructor(uiController) {
        this.ui = uiController;
        this.currentZoom = 1.0;
    }

    async render(file) {
        try {
            // console.log('📝 开始渲染DOCX文档');
            
            if (typeof mammoth === 'undefined') {
                throw new Error('Mammoth.js库未加载');
            }
            
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
            
            this._setupViewer(result.value);
            this._setupControls();
            this.ui.hideLoading();
            
            if (result.messages.length > 0) {
                console.warn('DOCX转换警告:', result.messages);
                this.ui.showWarnings(result.messages);
            }
            
            // console.log('✅ DOCX文档渲染完成');
        } catch (error) {
            throw new Error(`DOCX处理失败: ${error.message}`);
        }
    }
    
    _setupViewer(htmlContent) {
        const container = document.getElementById('viewerContainer');
        container.innerHTML = `
            <div class="docx-container" id="docxContent">
                ${htmlContent}
            </div>
        `;
    }

    _setupControls() {
        this.ui.showControls('docx');
        
        document.getElementById('docxZoomOut').onclick = () => this.zoomOut();
        document.getElementById('docxZoomIn').onclick = () => this.zoomIn();
        document.getElementById('docxFitWidth').onclick = () => this.fitWidth();
        
        this._updateZoomInfo();
    }

    _updateZoomInfo() {
        document.getElementById('docxZoomLevel').textContent = `${Math.round(this.currentZoom * 100)}%`;
    }

    _applyZoom() {
        const docxContent = document.getElementById('docxContent');
        if (docxContent) {
            docxContent.style.transform = `scale(${this.currentZoom})`;
            docxContent.style.transformOrigin = 'top center';
        }
        this._updateZoomInfo();
    }

    zoomIn() {
        this.currentZoom = Math.min(this.currentZoom + 0.1, 2.0);
        this._applyZoom();
    }

    zoomOut() {
        this.currentZoom = Math.max(this.currentZoom - 0.1, 0.5);
        this._applyZoom();
    }

    fitWidth() {
        this.currentZoom = 1.0;
        this._applyZoom();
    }
}

class PPTXRenderer {
    constructor(uiController) {
        this.ui = uiController;
    }

    async render(file) {
        try {
            // console.log('📊 开始渲染PPTX文档');
            
            if (typeof JSZip === 'undefined') {
                throw new Error('JSZip库未加载');
            }
            
            const container = document.getElementById('viewerContainer');
            
            if (typeof OfficeViewer !== 'undefined') {
                // console.log('🎯 使用OfficeViewer进行PPTX预览');
                const officeViewer = new OfficeViewer();
                await officeViewer.renderPPTXWithOfficeViewer(file, container);
                this.ui.hideLoading();
                return;
            }
            
            throw new Error('OfficeViewer库未加载，无法预览PPTX文件');
        } catch (error) {
            throw new Error(`PPTX处理失败: ${error.message}`);
        }
    }
    

}

class ExcelRenderer {
    constructor(uiController) {
        this.ui = uiController;
    }

    async render(file) {
        try {
            // console.log('📈 开始渲染Excel文档');
            
            if (typeof XLSX === 'undefined') {
                throw new Error('SheetJS (XLSX) 库未加载');
            }
            
            if (typeof ExcelViewer === 'undefined') {
                throw new Error('ExcelViewer未加载');
            }
            
        const container = document.getElementById('viewerContainer');
            if (!container) {
                throw new Error('viewerContainer元素未找到');
            }
            
            // console.log('✅ 开始创建ExcelViewer实例');
            const excelViewer = new ExcelViewer();
            
            // console.log('✅ 开始渲染Excel文件');
            await excelViewer.renderExcel(file, container);
            
            this.ui.hideLoading();
            this.ui.showControls('excel');
            
            // console.log('✅ Excel文档渲染完成');
        } catch (error) {
            console.error('❌ Excel处理详细错误:', error);
            throw error;
        }
    }
}

// ====================================================================================
// 文件处理器：负责文件验证和分发处理
// ====================================================================================
class FileHandler {
    constructor(uiController) {
        this.ui = uiController;
        this.currentFile = null;
        
        // 初始化各种渲染器
        this.pdfRenderer = new PDFRenderer(uiController);
        this.docxRenderer = new DOCXRenderer(uiController);
        this.pptxRenderer = new PPTXRenderer(uiController);
        this.excelRenderer = new ExcelRenderer(uiController);
    }

    /**
     * 处理用户选择的文件
     * @param {File} file - 用户选择的文件对象
     */
    async handleFile(file) {
        // console.log('📁 开始处理文件:', file.name);
        
        // 文件大小检查（50MB限制）
        const maxSize = 50 * 1024 * 1024; // 50MB
        if (file.size > maxSize) {
            this.ui.showError('文件过大，请选择小于50MB的文件');
            return;
        }
        
        // 获取文件扩展名
        const extension = DocumentUtils.getFileExtension(file.name);
        
        // 检查是否为支持的格式
        if (!DocumentUtils.isSupportedFormat(extension)) {
            this._handleUnsupportedFormat(extension);
            return;
        }
        
        // 保存当前文件引用
        this.currentFile = file;
        
        // 显示文件信息和加载状态
        this.ui.showFileInfo(file);
        this.ui.showLoading();
        
        try {
            // 根据文件类型分发给对应的渲染器处理
            await this._dispatchToRenderer(extension, file);
            // console.log('✅ 文件处理完成');
        } catch (error) {
            console.error('❌ 文件处理错误:', error);
            this.ui.showError(`文件处理失败: ${error.message}`);
        }
    }

    /**
     * 根据文件扩展名分发给对应的渲染器
     * @param {string} extension - 文件扩展名
     * @param {File} file - 文件对象
     * @private
     */
    async _dispatchToRenderer(extension, file) {
        switch (extension) {
            case 'pdf':
                await this.pdfRenderer.render(file);
                break;
            case 'docx':
                await this.docxRenderer.render(file);
                break;
            case 'pptx':
                await this.pptxRenderer.render(file);
                break;
            case 'xlsx':
            case 'xls':
            case 'xlsm':
            case 'xlsb':
            case 'csv':
                await this.excelRenderer.render(file);
                break;
            default:
                throw new Error(`不支持的文件格式: .${extension}`);
        }
    }

    /**
     * 处理不支持的文件格式
     * @param {string} extension - 文件扩展名
     * @private
     */
    _handleUnsupportedFormat(extension) {
        let message;
        switch (extension) {
            case 'doc':
                message = '暂不支持.doc格式，请转换为.docx格式';
                break;
            case 'ppt':
                message = '暂不支持.ppt格式，请转换为.pptx格式';
                break;
            default:
                message = `不支持的文件格式: .${extension}`;
        }
        this.ui.showError(message);
    }

    /**
     * 获取当前处理的文件
     * @returns {File|null} 当前文件对象
     */
    getCurrentFile() {
        return this.currentFile;
    }
}

// ====================================================================================
// 事件管理器：负责处理用户交互事件
// ====================================================================================
class EventManager {
    constructor(fileHandler) {
        this.fileHandler = fileHandler;
        this.eventsInitialized = false;
        this.fileChangeHandler = null;
    }

    /**
     * 初始化所有事件监听器
     */
    setupEventListeners() {
        // 防止重复绑定事件
        if (this.eventsInitialized) {
            // console.log('⚠️ 事件已初始化，跳过重复绑定');
            return;
        }
        
        // console.log('🎯 开始设置事件监听器');
        
        // 设置文件选择事件
        this._setupFileEvents();
        
        // 设置键盘快捷键
        this._setupKeyboardShortcuts();
        
        // 标记事件已初始化
        this.eventsInitialized = true;
        // console.log('✅ 事件监听器初始化完成');
    }

    /**
     * 设置文件选择相关事件
     * @private
     */
    _setupFileEvents() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        
        // 先移除可能存在的旧事件监听器
        if (this.fileChangeHandler) {
            fileInput.removeEventListener('change', this.fileChangeHandler);
        }
        
        // 创建文件选择事件处理函数
        this.fileChangeHandler = (e) => {
            // console.log('📁 文件选择事件触发');
            if (e.target.files.length > 0) {
                this.fileHandler.handleFile(e.target.files[0]);
            }
        };
        
        // 绑定文件选择事件
        fileInput.addEventListener('change', this.fileChangeHandler);
        
        // 设置点击上传区域事件
        // uploadArea.addEventListener('click', () => {
        //     fileInput.click();
        // });
    }


    /**
     * 设置键盘快捷键
     * @private
     */
    _setupKeyboardShortcuts() {
        document.addEventListener('keydown', (e) => {
            const currentFile = this.fileHandler.getCurrentFile();
            if (!currentFile) return;
            
            const extension = DocumentUtils.getFileExtension(currentFile.name);
            
            switch (e.key) {
                case 'ArrowLeft':
                    e.preventDefault();
                    this._handleLeftArrow(extension);
                    break;
                case 'ArrowRight':
                    e.preventDefault();
                    this._handleRightArrow(extension);
                    break;
                case 'Escape':
                    // 退出全屏
                    if (document.fullscreenElement) {
                        document.exitFullscreen();
                    }
                    break;
            }
        });
    }

    /**
     * 处理左方向键
     * @param {string} extension - 文件扩展名
     * @private
     */
    _handleLeftArrow(extension) {
        if (extension === 'pdf') {
            this.fileHandler.pdfRenderer.previousPage();
        }
        // PPTX导航功能由office-viewer.js内部处理
    }

    /**
     * 处理右方向键
     * @param {string} extension - 文件扩展名
     * @private
     */
    _handleRightArrow(extension) {
        if (extension === 'pdf') {
            this.fileHandler.pdfRenderer.nextPage();
        }
        // PPTX导航功能由office-viewer.js内部处理
    }

    /**
     * 清理事件监听器
     */
    cleanup() {
        const fileInput = document.getElementById('fileInput');
        if (fileInput && this.fileChangeHandler) {
            fileInput.removeEventListener('change', this.fileChangeHandler);
            this.fileChangeHandler = null;
        }
        this.eventsInitialized = false;
        // console.log('🧹 事件监听器已清理');
    }
}

// ====================================================================================
// 主文档查看器：应用程序的主控制器
// ====================================================================================
class OfflineDocumentViewer {
    constructor() {
        // console.log('🚀 初始化离线文档预览器');
        
        // 初始化各个组件
        this.uiController = new UIController();
        this.fileHandler = new FileHandler(this.uiController);
        this.eventManager = new EventManager(this.fileHandler);
        
        // 初始化应用
        this._init();
    }

    /**
     * 初始化应用程序
     * @private
     */
    _init() {
        // 设置PDF.js工作线程
        this._setupPDFWorker();
        
        // 设置事件监听器
        this.eventManager.setupEventListeners();
        
        // console.log('✅ 离线文档预览器初始化完成');
    }

    /**
     * 设置PDF.js工作线程
     * @private
     */
    _setupPDFWorker() {
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.js';
            // console.log('✅ PDF.js工作线程设置完成');
        }
    }

    /**
     * 清理资源
     */
    cleanup() {
        this.eventManager.cleanup();
        // console.log('🧹 文档查看器资源已清理');
    }
}

// ====================================================================================
// 应用程序初始化：页面加载完成后的初始化逻辑
// ====================================================================================

/**
 * 检查必要库文件是否已加载
 * @returns {Array} 缺失的库文件列表
 */
function checkRequiredLibraries() {
    // 检测是否启用了动态资源加载
    const hasDynamicLoading = typeof window.resourceManager !== 'undefined';
    
    if (hasDynamicLoading) {
        // 动态加载模式：只检查基础依赖
        const requiredLibs = [
            { name: 'jQuery', check: () => typeof $ !== 'undefined' }
        ];
        console.log('🎛️ 检测到动态资源管理器，启用按需加载模式');
        return requiredLibs.filter(lib => !lib.check()).map(lib => lib.name);
    } else {
        // 传统静态加载模式：检查所有库
        const requiredLibs = [
            { name: 'jQuery', check: () => typeof $ !== 'undefined' },
            { name: 'JSZip', check: () => typeof JSZip !== 'undefined' },
            { name: 'PDF.js', check: () => typeof pdfjsLib !== 'undefined' },
            { name: 'Mammoth.js', check: () => typeof mammoth !== 'undefined' }
        ];
        console.log('📚 使用传统静态加载模式');
        return requiredLibs.filter(lib => !lib.check()).map(lib => lib.name);
    }
}

/**
 * 动态加载Excel相关库和样式
 */
async function loadExcelSupport() {
    try {
        // console.log('🔄 正在加载Excel支持库...');
        
        // 加载SheetJS库（如果未加载）
        if (typeof XLSX === 'undefined') {
            await DocumentUtils.loadScript('lib/xlsx.full.min.js');
            // console.log('✅ SheetJS 库加载完成');
        }
        
        // 加载ExcelViewer（如果未加载）
        if (typeof ExcelViewer === 'undefined') {
            await DocumentUtils.loadScript('lib/excel-viewer.js');
            // console.log('✅ ExcelViewer 加载完成');
        }
        
        // 加载Excel样式
        await DocumentUtils.loadCSS('css/excel-styles.css');
        // console.log('✅ Excel样式加载完成');
        
        return true;
    } catch (error) {
        console.warn('⚠️ Excel库加载失败，Excel功能将不可用:', error);
        return false;
    }
}

/**
 * 动态添加Excel格式支持标识到界面
 */
function addExcelFormatBadges() {
        const formatContainer = document.querySelector('.supported-formats');
        if (formatContainer && !formatContainer.querySelector('.format-badge.xlsx')) {
        
        // 添加XLSX格式标识
            const xlsxBadge = document.createElement('span');
            xlsxBadge.className = 'format-badge xlsx';
            xlsxBadge.textContent = 'XLSX';
            xlsxBadge.style.cssText = `
                background: #28a745;
                color: white;
                padding: 4px 8px;
                border-radius: 12px;
                font-size: 12px;
                font-weight: bold;
                margin: 0 4px;
                display: inline-block;
            `;
            formatContainer.appendChild(xlsxBadge);
            
        // 添加CSV格式标识
            const csvBadge = document.createElement('span');
            csvBadge.className = 'format-badge csv';
            csvBadge.textContent = 'CSV';
            csvBadge.style.cssText = xlsxBadge.style.cssText.replace('#28a745', '#17a2b8');
            formatContainer.appendChild(csvBadge);
        
        // console.log('✅ Excel格式标识已添加到界面');
    }
}

/**
 * 显示初始化失败信息
 * @param {Array} missingLibs - 缺失的库文件列表
 */
function showInitializationError(missingLibs) {
    console.error('缺少必要的库:', missingLibs);
    document.getElementById('viewerContainer').innerHTML = `
        <div class="error-message">
            <h3>❌ 初始化失败</h3>
            <p>缺少必要的库文件: ${missingLibs.join(', ')}</p>
            <p>请运行 <code>npm run build</code> 下载依赖文件</p>
        </div>
    `;
}

// ====================================================================================
// 页面加载完成后的初始化入口点
// ====================================================================================
document.addEventListener('DOMContentLoaded', async () => {
    // console.log('📄 页面加载完成，开始初始化文档查看器');
    
    // 防止重复初始化
    if (window.documentViewerInitialized) {
        // console.log('⚠️ 文档查看器已初始化，跳过重复初始化');
        return;
    }
    
    // 标记正在初始化
    window.documentViewerInitialized = true;
    
    try {
        // 1. 检查必要的库是否加载
        const missingLibs = checkRequiredLibraries();
        
        if (missingLibs.length > 0) {
            showInitializationError(missingLibs);
            return;
        }
        
        // 2. 处理Excel支持（动态加载模式下跳过预加载）
        let excelSupported = false;
        const hasDynamicLoading = typeof window.resourceManager !== 'undefined';
        
        if (hasDynamicLoading) {
            // 动态加载模式：Excel将按需加载，这里标记为可用
            excelSupported = true;
            console.log('📊 Excel功能将按需加载');
        } else {
            // 传统模式：预加载Excel支持
            excelSupported = await loadExcelSupport();
        }
        
        // 3. 清理可能存在的旧实例
        if (window.documentViewer) {
            // console.log('⚠️ 发现旧的文档查看器实例，进行清理');
            window.documentViewer.cleanup();
        }
        
        // 4. 创建新的文档查看器实例
        window.documentViewer = new OfflineDocumentViewer();
        
        // 5. 如果Excel支持可用，添加格式标识
        if (excelSupported) {
            if (hasDynamicLoading) {
                // 动态加载模式：直接添加格式标识
                addExcelFormatBadges();
                console.log('✅ Excel预览功能（按需加载）已启用');
            } else if (typeof XLSX !== 'undefined' && typeof ExcelViewer !== 'undefined') {
                // 传统模式：检查库是否已加载
                addExcelFormatBadges();
                console.log('✅ Excel预览功能（预加载）已启用');
    } else {
        console.warn('⚠️ Excel预览功能不可用');
            }
        } else {
            console.warn('⚠️ Excel预览功能不可用');
        }
        
        // console.log('🎉 离线文档预览器初始化完成！支持的格式: PDF, DOCX, PPTX' + (excelSupported ? ', XLSX, CSV' : ''));
        
    } catch (error) {
        console.error('❌ 文档查看器初始化失败:', error);
        showInitializationError(['初始化过程']);
    }
});
