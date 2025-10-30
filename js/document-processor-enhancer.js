/**
 * 文档处理增强器
 * 在不修改现有document-viewer.js的情况下，为文件处理添加动态资源加载功能
 * 
 * @author AI Assistant
 * @version 1.0.0
 */

class DocumentProcessorEnhancer {
    constructor() {
        this.originalFileHandler = null;
        this.isEnhanced = false;
        this.resourceManager = null;
        this.init();
    }

    /**
     * 初始化增强器
     */
    init() {
        // 确保资源管理器已加载
        if (typeof window.resourceManager === 'undefined') {
            console.error('❌ 资源管理器未找到，无法启用动态加载');
            return;
        }
        
        this.resourceManager = window.resourceManager;
        
        // 等待文档查看器初始化完成后进行增强
        this.waitForDocumentViewer().then(() => {
            this.enhanceDocumentViewer();
        });
    }

    /**
     * 等待文档查看器初始化
     * @returns {Promise<void>}
     */
    async waitForDocumentViewer() {
        return new Promise((resolve) => {
            const checkViewer = () => {
                if (window.documentViewer && window.documentViewer.fileHandler) {
                    resolve();
                } else {
                    setTimeout(checkViewer, 100);
                }
            };
            checkViewer();
        });
    }

    /**
     * 增强文档查看器，添加动态资源加载
     */
    enhanceDocumentViewer() {
        if (this.isEnhanced) return;

        try {
            const fileHandler = window.documentViewer.fileHandler;
            if (!fileHandler || !fileHandler.handleFile) {
                console.warn('⚠️ 文件处理器不可用，跳过增强');
                return;
            }

            // 保存原始的handleFile方法
            this.originalFileHandler = fileHandler.handleFile.bind(fileHandler);

            // 创建增强版的handleFile方法
            fileHandler.handleFile = async (file) => {
                console.log('🚀 启用动态资源加载处理文件:', file.name);
                
                try {
                    // 显示加载状态
                    this.showResourceLoadingStatus(true);
                    
                    // 检测文件类型
                    const fileType = this.resourceManager.detectFileType(file);
                    console.log('📋 检测到文件类型:', fileType);
                    
                    // 检查当前资源状态
                    console.log('🔍 加载前资源状态:');
                    if (fileType === 'xlsx' || fileType === 'csv') {
                        console.log('  - XLSX:', typeof XLSX);
                        console.log('  - ExcelViewer:', typeof ExcelViewer);
                    }
                    
                    // 根据文件类型加载必需资源
                    const loadSuccess = await this.resourceManager.loadResourcesForFile(file);
                    
                    if (!loadSuccess) {
                        throw new Error('必需资源加载失败');
                    }
                    
                    // 再次检查资源状态
                    console.log('✅ 加载后资源状态:');
                    if (fileType === 'xlsx' || fileType === 'csv') {
                        console.log('  - XLSX:', typeof XLSX);
                        console.log('  - ExcelViewer:', typeof ExcelViewer);
                        
                        // 额外确认ExcelViewer可用
                        if (typeof ExcelViewer !== 'undefined') {
                            console.log('  - ExcelViewer实例化测试:', new ExcelViewer() instanceof Object);
                        }
                    }
                    
                    // 隐藏加载状态
                    this.showResourceLoadingStatus(false);
                    
                    console.log('🎯 调用原始处理器...');
                    // 调用原始的文件处理逻辑
                    return await this.originalFileHandler(file);
                    
                } catch (error) {
                    console.error('❌ 文件处理失败:', error);
                    this.showResourceLoadingStatus(false);
                    
                    // 显示用户友好的错误信息
                    this.showResourceLoadError(error.message);
                    throw error;
                }
            };

            this.isEnhanced = true;
            console.log('✅ 文档处理器已增强，启用动态资源加载');

        } catch (error) {
            console.error('❌ 文档处理器增强失败:', error);
        }
    }

    /**
     * 显示资源加载状态
     * @param {boolean} loading - 是否正在加载
     */
    showResourceLoadingStatus(loading) {
        const loadingOverlay = document.getElementById('loadingOverlay');
        const loadingText = document.querySelector('.loading-text');
        
        if (loadingOverlay) {
            if (loading) {
                loadingOverlay.style.display = 'flex';
                if (loadingText) {
                    loadingText.textContent = '正在加载必需资源...';
                }
            } else {
                // 延迟隐藏，让用户看到加载完成
                setTimeout(() => {
                    if (loadingText) {
                        loadingText.textContent = '正在处理文档...';
                    }
                }, 300);
            }
        }
    }

    /**
     * 显示资源加载错误
     * @param {string} errorMessage - 错误信息
     */
    showResourceLoadError(errorMessage) {
        const viewerContainer = document.getElementById('viewerContainer');
        if (viewerContainer && window.documentViewer && window.documentViewer.uiController) {
            window.documentViewer.uiController.showError(`资源加载失败: ${errorMessage}`);
        }
    }

    /**
     * 预加载用户可能使用的资源
     * @param {Array<string>} fileTypes - 要预加载的文件类型
     * @param {boolean} force - 是否强制预加载
     */
    async preloadLikelyResources(fileTypes = [], force = false) {
        if (!this.resourceManager || fileTypes.length === 0) return;
        
        // 检查用户是否禁用了预加载
        const preloadDisabled = localStorage.getItem('documentViewer_disablePreload') === 'true';
        if (preloadDisabled && !force) {
            console.log('⏸️ 用户已禁用预加载功能');
            return;
        }
        
        console.log('🎯 预加载常用资源中...', fileTypes);
        await this.resourceManager.preloadResources(fileTypes);
    }

    /**
     * 获取增强器状态
     * @returns {Object} 状态信息
     */
    getStatus() {
        return {
            isEnhanced: this.isEnhanced,
            hasResourceManager: !!this.resourceManager,
            hasOriginalHandler: !!this.originalFileHandler
        };
    }

    /**
     * 重置增强器（调试用）
     */
    reset() {
        if (this.isEnhanced && this.originalFileHandler && window.documentViewer) {
            window.documentViewer.fileHandler.handleFile = this.originalFileHandler;
            this.isEnhanced = false;
            console.log('🔄 文档处理器增强已重置');
        }
    }
}

/**
 * 智能资源预加载策略
 * 根据用户行为模式预加载资源
 */
class SmartPreloadStrategy {
    constructor(enhancer) {
        this.enhancer = enhancer;
        this.userPreferences = this.loadUserPreferences();
        this.setupIntelligentPreload();
    }

    /**
     * 加载用户偏好（从localStorage）
     * @returns {Object} 用户偏好数据
     */
    loadUserPreferences() {
        try {
            const prefs = localStorage.getItem('documentViewer_userPrefs');
            return prefs ? JSON.parse(prefs) : { frequentTypes: [] };
        } catch {
            return { frequentTypes: [] };
        }
    }

    /**
     * 保存用户偏好
     * @param {Object} preferences - 偏好数据
     */
    saveUserPreferences(preferences) {
        try {
            localStorage.setItem('documentViewer_userPrefs', JSON.stringify(preferences));
        } catch (error) {
            console.warn('⚠️ 无法保存用户偏好:', error);
        }
    }

    /**
     * 记录文件类型使用
     * @param {string} fileType - 文件类型
     */
    recordFileTypeUsage(fileType) {
        if (!this.userPreferences.frequentTypes) {
            this.userPreferences.frequentTypes = [];
        }
        
        this.userPreferences.frequentTypes.unshift(fileType);
        
        // 保持最近10次记录
        if (this.userPreferences.frequentTypes.length > 10) {
            this.userPreferences.frequentTypes = this.userPreferences.frequentTypes.slice(0, 10);
        }
        
        this.saveUserPreferences(this.userPreferences);
    }

    /**
     * 设置智能预加载
     */
    setupIntelligentPreload() {
        // 页面空闲时预加载常用资源
        if ('requestIdleCallback' in window) {
            requestIdleCallback(() => {
                this.performIntelligentPreload();
            });
        } else {
            // 降级到setTimeout
            setTimeout(() => {
                this.performIntelligentPreload();
            }, 2000);
        }
    }

    /**
     * 执行智能预加载
     */
    async performIntelligentPreload() {
        // 检查用户预加载设置，默认禁用预加载（完全按需加载）
        const preloadSetting = localStorage.getItem('documentViewer_disablePreload');
        const preloadDisabled = preloadSetting === null ? true : preloadSetting === 'true';
        
        if (preloadDisabled) {
            console.log('⏸️ 预加载已禁用，启用完全按需加载模式');
            // 如果是第一次访问，设置默认禁用预加载
            if (preloadSetting === null) {
                localStorage.setItem('documentViewer_disablePreload', 'true');
            }
            return;
        }
        
        const { frequentTypes = [] } = this.userPreferences;
        
        // 分析用户最常使用的文件类型
        const typeFrequency = {};
        frequentTypes.forEach(type => {
            typeFrequency[type] = (typeFrequency[type] || 0) + 1;
        });
        
        // 按频率排序，预加载前2种
        const sortedTypes = Object.entries(typeFrequency)
            .sort(([,a], [,b]) => b - a)
            .slice(0, 2)
            .map(([type]) => type);
        
        if (sortedTypes.length > 0) {
            console.log('🎯 基于用户偏好预加载资源:', sortedTypes);
            await this.enhancer.preloadLikelyResources(sortedTypes);
        } else {
            // 首次使用，完全按需加载，不预加载任何资源
            console.log('💡 首次使用，启用完全按需加载模式');
        }
    }

    /**
     * 拦截文件处理，记录使用模式
     */
    interceptFileProcessing() {
        if (!window.resourceManager) return;

        const originalLoadForFile = window.resourceManager.loadResourcesForFile.bind(window.resourceManager);
        
        window.resourceManager.loadResourcesForFile = async (file) => {
            const fileType = window.resourceManager.detectFileType(file);
            
            // 记录使用模式
            if (fileType !== 'unknown') {
                this.recordFileTypeUsage(fileType);
            }
            
            return await originalLoadForFile(file);
        };
    }
}

// 初始化增强器
if (typeof window !== 'undefined') {
    // 确保DOM完全加载后初始化
    document.addEventListener('DOMContentLoaded', () => {
        window.documentEnhancer = new DocumentProcessorEnhancer();
        window.smartPreloadStrategy = new SmartPreloadStrategy(window.documentEnhancer);
        
        // 设置使用模式拦截
        setTimeout(() => {
            window.smartPreloadStrategy.interceptFileProcessing();
        }, 1000);
        
        // 提供快速重置功能（调试用）
        window.resetToOnDemandLoading = function() {
            localStorage.removeItem('documentViewer_userPrefs');
            localStorage.setItem('documentViewer_disablePreload', 'true');
            console.log('🔄 已重置为完全按需加载模式');
            console.log('💡 请刷新页面查看效果');
        };
        
        console.log('🎛️ 文档处理增强器已启动');
        console.log('💡 如需重置为完全按需加载，请在控制台执行：resetToOnDemandLoading()');
    });
}
