/**
 * 动态资源管理器
 * 根据文档类型按需加载相应的第三方库，避免不必要的资源浪费
 * 
 * @author AI Assistant
 * @version 1.0.0
 */

class DynamicResourceManager {
    constructor() {
        this.loadedResources = new Set();
        this.loadingPromises = new Map();
        this.resourceConfig = this.initResourceConfig();
    }

    /**
     * 初始化资源配置表
     * @returns {Object} 资源配置映射
     */
    initResourceConfig() {
        return {
            // PDF 相关资源
            'pdf': {
                scripts: ['lib/pdf.min.js'],
                styles: [],
                dependencies: [], // PDF不依赖JSZip
                checker: () => typeof pdfjsLib !== 'undefined'
            },
            
            // DOCX 相关资源  
            'docx': {
                scripts: ['lib/jszip.min.js', 'lib/mammoth.min.js'],
                styles: [],
                dependencies: ['jszip'],
                checker: () => typeof mammoth !== 'undefined' && typeof JSZip !== 'undefined'
            },
            
            // PPTX 相关资源
            'pptx': {
                scripts: ['lib/jszip.min.js', 'lib/office-viewer.js'],
                styles: ['css/pptx-styles.css'],
                dependencies: ['jszip'],
                checker: () => typeof OfficeViewer !== 'undefined' && typeof JSZip !== 'undefined'
            },
            
            // Excel 相关资源 (复用现有逻辑)
            'xlsx': {
                scripts: ['lib/xlsx.full.min.js', 'lib/excel-viewer.js'],
                styles: ['css/excel-styles.css'],
                dependencies: [],
                checker: () => typeof XLSX !== 'undefined' && typeof ExcelViewer !== 'undefined'
            },
            
            'csv': {
                scripts: ['lib/xlsx.full.min.js', 'lib/excel-viewer.js'],
                styles: ['css/excel-styles.css'],
                dependencies: [],
                checker: () => typeof XLSX !== 'undefined' && typeof ExcelViewer !== 'undefined'
            },
            
            // 共享依赖
            'jszip': {
                scripts: ['lib/jszip.min.js'],
                styles: [],
                dependencies: [],
                checker: () => typeof JSZip !== 'undefined'
            }
        };
    }

    /**
     * 根据文件类型加载必需的资源
     * @param {string} fileType - 文件类型 (pdf, docx, pptx, xlsx, csv)
     * @returns {Promise<boolean>} 是否加载成功
     */
    async loadResourcesForFileType(fileType) {
        console.log(`🔄 开始为 ${fileType.toUpperCase()} 加载必需资源...`);
        
        try {
            const config = this.resourceConfig[fileType.toLowerCase()];
            if (!config) {
                console.warn(`⚠️ 未找到 ${fileType} 的资源配置`);
                return false;
            }

            // 检查是否已经加载
            console.log(`🔍 检查 ${fileType.toUpperCase()} 资源状态...`);
            if (config.checker()) {
                console.log(`✅ ${fileType.toUpperCase()} 资源已加载，跳过`);
                return true;
            }
            
            console.log(`📋 ${fileType.toUpperCase()} 需要加载的资源:`, config.scripts.concat(config.styles));

            // 先加载依赖
            for (const dep of config.dependencies) {
                await this.loadDependency(dep);
            }

            // 顺序加载脚本文件（保证依赖关系）
            for (const script of config.scripts) {
                await this.loadScript(script);
            }
            
            // 并行加载样式文件
            const styleTasks = [];
            for (const style of config.styles) {
                styleTasks.push(this.loadCSS(style));
            }
            
            if (styleTasks.length > 0) {
                await Promise.all(styleTasks);
            }

            // 等待脚本执行（给浏览器时间执行脚本）
            await new Promise(resolve => setTimeout(resolve, 100));
            
            // 对于Excel文件，额外等待确保XLSX和ExcelViewer都完全可用
            if (fileType.toLowerCase() === 'xlsx' || fileType.toLowerCase() === 'csv') {
                let attempts = 0;
                const maxAttempts = 20; // 最多等待2秒
                
                while (attempts < maxAttempts && (typeof XLSX === 'undefined' || typeof ExcelViewer === 'undefined')) {
                    await new Promise(resolve => setTimeout(resolve, 100));
                    attempts++;
                    console.log(`⏳ 等待Excel库完全加载... (${attempts}/${maxAttempts})`);
                }
                
                if (typeof XLSX === 'undefined' || typeof ExcelViewer === 'undefined') {
                    console.error('❌ Excel库等待超时，XLSX:', typeof XLSX, 'ExcelViewer:', typeof ExcelViewer);
                } else {
                    console.log('✅ Excel库完全加载确认');
                }
            }

            // 验证加载结果
            console.log(`🔎 验证 ${fileType.toUpperCase()} 资源加载状态...`);
            const verificationResult = this.detailedVerification(fileType, config);
            
            if (verificationResult.success) {
                console.log(`✅ ${fileType.toUpperCase()} 资源加载完成`);
                return true;
            } else {
                // 再次等待并重试验证
                console.warn(`⚠️ ${fileType} 首次验证失败: ${verificationResult.error}，等待重试...`);
                await new Promise(resolve => setTimeout(resolve, 200));
                
                const retryResult = this.detailedVerification(fileType, config);
                if (retryResult.success) {
                    console.log(`✅ ${fileType.toUpperCase()} 资源加载完成（重试成功）`);
                    return true;
                } else {
                    throw new Error(`${fileType} 资源加载验证失败: ${retryResult.error}`);
                }
            }

        } catch (error) {
            console.error(`❌ ${fileType.toUpperCase()} 资源加载失败:`, error);
            return false;
        }
    }

    /**
     * 详细验证资源加载状态
     * @param {string} fileType - 文件类型
     * @param {Object} config - 资源配置
     * @returns {Object} 验证结果
     */
    detailedVerification(fileType, config) {
        try {
            const result = config.checker();
            if (result) {
                return { success: true };
            }
            
            // 详细检查各个依赖
            const errors = [];
            switch (fileType.toLowerCase()) {
                case 'xlsx':
                case 'csv':
                    if (typeof XLSX === 'undefined') {
                        errors.push('XLSX library not loaded');
                    }
                    if (typeof ExcelViewer === 'undefined') {
                        errors.push('ExcelViewer class not defined');
                    }
                    break;
                case 'pdf':
                    if (typeof pdfjsLib === 'undefined') {
                        errors.push('PDF.js library not loaded');
                    }
                    break;
                case 'docx':
                    if (typeof JSZip === 'undefined') {
                        errors.push('JSZip library not loaded');
                    }
                    if (typeof mammoth === 'undefined') {
                        errors.push('Mammoth.js library not loaded');
                    }
                    break;
                case 'pptx':
                    if (typeof JSZip === 'undefined') {
                        errors.push('JSZip library not loaded');
                    }
                    if (typeof OfficeViewer === 'undefined') {
                        errors.push('OfficeViewer class not defined');
                    }
                    break;
            }
            
            return { success: false, error: errors.join(', ') || 'Unknown verification error' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * 加载依赖资源
     * @param {string} depName - 依赖名称
     * @returns {Promise<void>}
     */
    async loadDependency(depName) {
        const config = this.resourceConfig[depName];
        if (!config) return;

        if (config.checker()) return; // 已加载

        // 加载依赖的脚本
        for (const script of config.scripts) {
            await this.loadScript(script);
        }
    }

    /**
     * 动态加载JavaScript文件
     * @param {string} src - 脚本路径
     * @returns {Promise<void>}
     */
    async loadScript(src) {
        // 避免重复加载
        if (this.loadedResources.has(src)) {
            return;
        }

        // 如果正在加载，返回现有Promise
        if (this.loadingPromises.has(src)) {
            return this.loadingPromises.get(src);
        }

        // 创建加载Promise
        const loadPromise = new Promise((resolve, reject) => {
            // 先检查脚本是否已经存在于页面中
            const existingScript = document.querySelector(`script[src="${src}"]`);
            if (existingScript) {
                this.loadedResources.add(src);
                console.log(`📦 脚本已存在于页面: ${src}`);
                resolve();
                return;
            }
            
            console.log(`🔄 开始加载脚本: ${src}`);
            const script = document.createElement('script');
            script.src = src;
            script.async = false; // 改为同步，确保执行顺序
            
            script.onload = () => {
                this.loadedResources.add(src);
                console.log(`📦 脚本加载完成: ${src}`);
                
                // 额外验证：对于关键库，检查是否正确导出
                if (src.includes('xlsx.full.min.js')) {
                    console.log(`  ✅ XLSX加载后状态: ${typeof XLSX}`);
                }
                if (src.includes('excel-viewer.js')) {
                    console.log(`  ✅ ExcelViewer加载后状态: ${typeof ExcelViewer}`);
                }
                
                resolve();
            };
            
            script.onerror = (error) => {
                console.error(`❌ 脚本加载失败: ${src}`, error);
                reject(new Error(`Failed to load script: ${src}`));
            };
            
            document.head.appendChild(script);
        });

        this.loadingPromises.set(src, loadPromise);
        
        try {
            await loadPromise;
        } finally {
            this.loadingPromises.delete(src);
        }
    }

    /**
     * 动态加载CSS文件
     * @param {string} href - 样式表路径
     * @returns {Promise<void>}
     */
    async loadCSS(href) {
        // 避免重复加载
        if (this.loadedResources.has(href)) {
            return;
        }

        // 如果正在加载，返回现有Promise
        if (this.loadingPromises.has(href)) {
            return this.loadingPromises.get(href);
        }

        // 创建加载Promise
        const loadPromise = new Promise((resolve, reject) => {
            const link = document.createElement('link');
            link.rel = 'stylesheet';
            link.href = href;
            
            link.onload = () => {
                this.loadedResources.add(href);
                console.log(`🎨 样式加载完成: ${href}`);
                resolve();
            };
            
            link.onerror = () => {
                reject(new Error(`Failed to load CSS: ${href}`));
            };
            
            document.head.appendChild(link);
        });

        this.loadingPromises.set(href, loadPromise);
        
        try {
            await loadPromise;
        } finally {
            this.loadingPromises.delete(href);
        }
    }

    /**
     * 预加载常用资源（可选优化）
     * @param {Array<string>} fileTypes - 要预加载的文件类型
     * @returns {Promise<void>}
     */
    async preloadResources(fileTypes = []) {
        console.log('🚀 开始预加载资源...');
        
        const preloadTasks = fileTypes.map(type => 
            this.loadResourcesForFileType(type).catch(err => 
                console.warn(`预加载 ${type} 失败:`, err)
            )
        );
        
        await Promise.all(preloadTasks);
        console.log('🎯 资源预加载完成');
    }

    /**
     * 获取已加载资源统计
     * @returns {Object} 统计信息
     */
    getLoadedResourcesStats() {
        const stats = {
            totalLoaded: this.loadedResources.size,
            loadedResources: Array.from(this.loadedResources),
            loadingCount: this.loadingPromises.size
        };
        
        console.log('📊 资源加载统计:', stats);
        return stats;
    }

    /**
     * 检测文件类型
     * @param {File} file - 文件对象
     * @returns {string} 文件类型
     */
    detectFileType(file) {
        if (!file || !file.name) return 'unknown';
        
        const ext = file.name.toLowerCase().split('.').pop();
        const typeMap = {
            'pdf': 'pdf',
            'docx': 'docx', 'doc': 'docx',
            'pptx': 'pptx', 'ppt': 'pptx', 
            'xlsx': 'xlsx', 'xls': 'xlsx',
            'csv': 'csv'
        };
        
        return typeMap[ext] || 'unknown';
    }

    /**
     * 智能资源加载 - 组合检测和加载
     * @param {File} file - 文件对象
     * @returns {Promise<boolean>} 加载是否成功
     */
    async loadResourcesForFile(file) {
        const fileType = this.detectFileType(file);
        
        if (fileType === 'unknown') {
            console.warn('⚠️ 未知文件类型，跳过资源加载');
            return false;
        }
        
        return await this.loadResourcesForFileType(fileType);
    }
}

// 创建全局实例
if (typeof window !== 'undefined') {
    window.DynamicResourceManager = DynamicResourceManager;
    window.resourceManager = new DynamicResourceManager();
    
    console.log('🎛️ 动态资源管理器已初始化');
}

// Node.js 环境支持
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DynamicResourceManager;
}
