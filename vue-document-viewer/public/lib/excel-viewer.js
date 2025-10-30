/**
 * Excel查看器 - 基于SheetJS实现
 * 支持XLSX、XLS、CSV等格式的在线预览
 */

class ExcelViewer {
    constructor() {
        this.currentFile = null;
        this.workbook = null;
        this.currentSheetIndex = 0;
        this.sheetNames = [];
    }

    /**
     * 渲染Excel文件
     * @param {File} file - Excel文件对象
     * @param {HTMLElement} container - 容器元素
     */
    async renderExcel(file, container) {
        try {
            
            // 显示加载状态
            this.showLoading(container);
            
            // 检查XLSX是否加载
            if (typeof XLSX === 'undefined') {
                throw new Error('SheetJS (XLSX) 库未加载');
            }
            
            // 根据文件类型选择不同的读取方式
            let workbookData;
            const fileName = file.name.toLowerCase();
            
            if (fileName.endsWith('.csv')) {
                // CSV文件：使用UTF-8编码读取
                workbookData = await this.readCSVWithUTF8(file);
                this.workbook = XLSX.read(workbookData, {
                    type: 'string',
                    cellDates: true,
                    cellNF: false,
                    cellText: false,
                    raw: false,
                    codepage: 65001  // UTF-8 codepage
                });
            } else {
                // Excel文件：使用arrayBuffer读取
                const arrayBuffer = await file.arrayBuffer();
                this.workbook = XLSX.read(arrayBuffer, {
                    type: 'array',
                    cellDates: true,
                    cellNF: false,
                    cellText: false,
                });
            }
            
            this.currentFile = file;
            this.sheetNames = this.workbook.SheetNames;
            this.currentSheetIndex = 0;
            
            
            // 渲染Excel查看器界面
            this.renderExcelViewer(container);
            
            // 确保CSS样式已加载
            this.ensureStylesLoaded();
            
        } catch (error) {
            console.error('❌ Excel渲染失败:', error);
            console.error('错误堆栈:', error.stack);
            this.showError(container, `Excel文件加载失败: ${error.message}`);
        }
    }

    /**
     * 渲染Excel查看器界面
     * @param {HTMLElement} container - 容器元素
     */
    renderExcelViewer(container) {
        console.log('🔄 开始渲染Excel查看器界面');
        console.log('当前文件:', this.currentFile);
        console.log('工作表名称:', this.sheetNames);
        console.log('工作簿对象:', this.workbook);
        
        try {
            const html = `
            <div class="excel-viewer">
                <!-- Excel文件信息 -->
                <div class="excel-header">
                    <div class="excel-info">
                        <h3>📊 ${this.currentFile.name}</h3>
                        <p>共 ${this.sheetNames.length} 个工作表 | 文件大小: ${this.formatFileSize(this.currentFile.size)}</p>
                    </div>
                    <div class="excel-controls">
                        <button class="btn btn-secondary" id="excelExportCSV">📄 导出CSV</button>
                        <button class="btn btn-secondary" id="excelExportJSON">📋 导出JSON</button>
                    </div>
                </div>

                <!-- 工作表选项卡 -->
                <div class="sheet-tabs">
                    ${this.renderSheetTabs()}
                </div>

                <!-- 表格显示区域 -->
                <div class="excel-content">
                    <div class="excel-toolbar">
                        <div class="sheet-info">
                            <span id="currentSheetInfo">${this.getCurrentSheetInfo()}</span>
                        </div>
                        <div class="view-controls">
                            <button class="btn btn-sm" id="excelFitWidth">适应宽度</button>
                            <button class="btn btn-sm" id="excelFullscreen">全屏</button>
                        </div>
                    </div>
                    <div class="excel-table-container" id="excelTableContainer">
                        ${this.renderCurrentSheet()}
                    </div>
                </div>
            </div>
        `;

            console.log('✅ HTML模板生成成功');
            container.innerHTML = html;
            console.log('✅ HTML内容已设置到容器');
            this.bindExcelEvents();
            console.log('✅ Excel事件绑定已启动');
            
        } catch (error) {
            console.error('❌ 渲染Excel查看器界面失败:', error);
            console.error('错误堆栈:', error.stack);
            this.showError(container, `界面渲染失败: ${error.message}`);
        }
    }

    /**
     * 渲染工作表选项卡
     */
    renderSheetTabs() {
        try {
            if (!this.sheetNames || this.sheetNames.length === 0) {
                return '<div class="no-sheets">没有可用的工作表</div>';
            }
            
            return this.sheetNames.map((name, index) => {
                const isActive = index === this.currentSheetIndex;
                return `
                    <button class="sheet-tab ${isActive ? 'active' : ''}" 
                            data-sheet-index="${index}"
                            title="${name || '未命名工作表'}">
                        ${this.escapeHtml(name || '未命名工作表')}
                    </button>
                `;
            }).join('');
        } catch (error) {
            console.error('❌ 渲染工作表选项卡失败:', error);
            return '<div class="sheet-tabs-error">选项卡渲染失败</div>';
        }
    }

    /**
     * 渲染当前工作表
     */
    renderCurrentSheet() {
        try {
            console.log('🔄 开始渲染当前工作表');
            console.log('当前工作表索引:', this.currentSheetIndex);
            console.log('工作表名称数组:', this.sheetNames);
            
            if (!this.sheetNames || this.sheetNames.length === 0) {
                console.warn('⚠️ 没有可用的工作表');
                return '<div class="excel-empty">没有可用的工作表</div>';
            }
            
            const sheetName = this.sheetNames[this.currentSheetIndex];
            console.log('当前工作表名称:', sheetName);
            
            if (!this.workbook || !this.workbook.Sheets) {
                console.warn('⚠️ 工作簿对象无效');
                return '<div class="excel-empty">工作簿数据无效</div>';
            }
            
            const worksheet = this.workbook.Sheets[sheetName];
            console.log('当前工作表对象:', worksheet);
            
            if (!worksheet || !worksheet['!ref']) {
                console.log('📄 工作表为空');
                return '<div class="excel-empty">此工作表为空</div>';
            }

            console.log('🔄 开始转换为HTML表格');
            // 转换为HTML表格，不显示行号
            const html = XLSX.utils.sheet_to_html(worksheet, {
                id: 'excel-table',
                editable: false,
                tableClass: 'excel-data-table'
            });
            
            // 移除行号列（第一列通常是行号）
            const cleanHtml = this.removeRowNumbers(html);

            console.log('✅ HTML表格转换成功');
            return `
                <div class="excel-sheet-wrapper">
                    ${cleanHtml}
                </div>
            `;
        } catch (error) {
            console.error('❌ 渲染工作表失败:', error);
            console.error('错误堆栈:', error.stack);
            return `<div class="excel-error">工作表渲染失败: ${error.message}</div>`;
        }
    }

    /**
     * 获取当前工作表信息
     */
    getCurrentSheetInfo() {
        try {
            if (!this.sheetNames || this.sheetNames.length === 0) {
                return '无工作表';
            }
            
            const sheetName = this.sheetNames[this.currentSheetIndex];
            if (!sheetName) {
                return '工作表名称无效';
            }
            
            if (!this.workbook || !this.workbook.Sheets) {
                return '工作簿无效';
            }
            
            const worksheet = this.workbook.Sheets[sheetName];
            
            if (!worksheet || !worksheet['!ref']) {
                return '空工作表';
            }

            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const rows = range.e.r - range.s.r + 1;
            const cols = range.e.c - range.s.c + 1;
            
            return `${rows} 行 × ${cols} 列`;
        } catch (error) {
            console.error('❌ 获取工作表信息失败:', error);
            return '信息获取失败';
        }
    }

    /**
     * 绑定Excel查看器事件
     */
    bindExcelEvents() {
        // 使用setTimeout确保DOM已经渲染完成
        setTimeout(() => {
            // 工作表切换
            const sheetTabs = document.querySelectorAll('.sheet-tab');
            sheetTabs.forEach(tab => {
                tab.addEventListener('click', (e) => {
                    const sheetIndex = parseInt(e.target.getAttribute('data-sheet-index'));
                    this.switchSheet(sheetIndex);
                });
            });

            // 导出功能
            const exportCsvBtn = document.getElementById('excelExportCSV');
            const exportJsonBtn = document.getElementById('excelExportJSON');
            
            if (exportCsvBtn) {
                exportCsvBtn.addEventListener('click', () => this.exportCurrentSheet('csv'));
            }
            
            if (exportJsonBtn) {
                exportJsonBtn.addEventListener('click', () => this.exportCurrentSheet('json'));
            }

            // 视图控制
            const fitWidthBtn = document.getElementById('excelFitWidth');
            const fullscreenBtn = document.getElementById('excelFullscreen');
            
            if (fitWidthBtn) {
                fitWidthBtn.addEventListener('click', () => this.fitTableWidth());
            }
            
            if (fullscreenBtn) {
                fullscreenBtn.addEventListener('click', () => this.toggleFullscreen());
            }
            
            console.log('✅ Excel事件绑定完成');
        }, 100);
    }

    /**
     * 切换工作表
     * @param {number} sheetIndex - 工作表索引
     */
    switchSheet(sheetIndex) {
        if (sheetIndex >= 0 && sheetIndex < this.sheetNames.length) {
            this.currentSheetIndex = sheetIndex;
            
            // 更新选项卡状态
            const tabs = document.querySelectorAll('.sheet-tab');
            tabs.forEach((tab, index) => {
                tab.classList.toggle('active', index === sheetIndex);
            });
            
            // 更新表格内容
            const container = document.getElementById('excelTableContainer');
            if (container) {
                container.innerHTML = this.renderCurrentSheet();
            }
            
            // 更新工作表信息
            const infoElement = document.getElementById('currentSheetInfo');
            if (infoElement) {
                infoElement.textContent = this.getCurrentSheetInfo();
            }
        }
    }

    /**
     * 导出当前工作表
     * @param {string} format - 导出格式 ('csv' 或 'json')
     */
    exportCurrentSheet(format) {
        try {
            const sheetName = this.sheetNames[this.currentSheetIndex];
            const worksheet = this.workbook.Sheets[sheetName];
            
            let data, mimeType, extension;
            
            if (format === 'csv') {
                data = XLSX.utils.sheet_to_csv(worksheet);
                mimeType = 'text/csv';
                extension = 'csv';
            } else if (format === 'json') {
                const jsonData = this.convertSheetToCleanJSON(worksheet);
                data = JSON.stringify(jsonData, null, 2);
                mimeType = 'application/json';
                extension = 'json';
            } else {
                throw new Error('不支持的导出格式');
            }
            
            // 创建下载链接
            const blob = new Blob([data], { type: mimeType });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            
            link.href = url;
            link.download = `${sheetName}.${extension}`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            
            console.log(`✅ 导出${format.toUpperCase()}成功:`, link.download);
            
        } catch (error) {
            console.error(`❌ 导出${format.toUpperCase()}失败:`, error);
            alert(`导出失败: ${error.message}`);
        }
    }

    /**
     * 适应表格宽度
     */
    fitTableWidth() {
        const table = document.getElementById('excel-table');
        if (table && table.style) {
            table.style.width = '100%';
            table.style.maxWidth = 'none';
            console.log('✅ 表格已适应容器宽度');
        } else {
            console.warn('⚠️ 未找到Excel表格元素');
        }
    }

    /**
     * 切换全屏模式
     */
    toggleFullscreen() {
        const container = document.querySelector('.excel-viewer');
        if (!container) return;
        
        if (!document.fullscreenElement) {
            container.requestFullscreen().then(() => {
                container.classList.add('fullscreen-mode');
                console.log('✅ 进入全屏模式');
            }).catch(err => {
                console.error('❌ 全屏模式失败:', err);
            });
        } else {
            document.exitFullscreen().then(() => {
                container.classList.remove('fullscreen-mode');
                console.log('✅ 退出全屏模式');
            });
        }
    }

    /**
     * 显示加载状态
     * @param {HTMLElement} container - 容器元素
     */
    showLoading(container) {
        container.innerHTML = `
            <div class="excel-loading">
                <div class="loading-spinner"></div>
                <h3>正在加载Excel文件...</h3>
                <p>请稍候，正在解析文件内容</p>
            </div>
        `;
    }

    /**
     * 显示错误信息
     * @param {HTMLElement} container - 容器元素
     * @param {string} message - 错误信息
     */
    showError(container, message) {
        container.innerHTML = `
            <div class="excel-error">
                <div class="error-icon">❌</div>
                <h3>Excel文件加载失败</h3>
                <p>${this.escapeHtml(message)}</p>
                <button class="btn btn-primary" onclick="location.reload()">重新加载</button>
            </div>
        `;
    }

    /**
     * 格式化文件大小
     * @param {number} bytes - 字节数
     * @returns {string} 格式化后的文件大小
     */
    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    /**
     * 移除HTML表格中的行号和序号
     * @param {string} html - 原始HTML字符串
     * @returns {string} 清理后的HTML字符串
     */
    removeRowNumbers(html) {
        try {
            // 创建临时DOM元素来处理HTML
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = html;
            
            const table = tempDiv.querySelector('table');
            if (!table) return html;
            
            const rows = table.querySelectorAll('tr');
            if (rows.length === 0) return html;
            
            // 移除可能的表格标题序号（第一行第一列如果是单独的数字）
            const firstRow = rows[0];
            if (firstRow) {
                const firstCell = firstRow.querySelector('td:first-child, th:first-child');
                if (firstCell) {
                    const text = firstCell.textContent.trim();
                    // 如果第一个单元格只包含数字（如 "1"），且没有其他有意义内容，则移除
                    if (/^\d+$/.test(text) && text.length <= 3) {
                        firstCell.remove();
                    }
                }
            }
            
            // 检查并移除连续的行号列
            let isFirstColumnRowNumbers = true;
            const checkRows = Math.min(rows.length, 5); // 检查前5行
            
            for (let i = 0; i < checkRows; i++) {
                const firstCell = rows[i].querySelector('td:first-child, th:first-child');
                if (firstCell) {
                    const text = firstCell.textContent.trim();
                    // 检查是否为连续的行号
                    if (!/^\d+$/.test(text) || parseInt(text) !== (i + 1)) {
                        isFirstColumnRowNumbers = false;
                        break;
                    }
                }
            }
            
            // 如果确认第一列是连续行号，则移除整列
            if (isFirstColumnRowNumbers) {
                rows.forEach(row => {
                    const firstCell = row.querySelector('td:first-child, th:first-child');
                    if (firstCell) {
                        firstCell.remove();
                    }
                });
            }
            
            return tempDiv.innerHTML;
            
        } catch (error) {
            return html;
        }
    }

    /**
     * 确保样式已加载
     */
    ensureStylesLoaded() {
        // 检查Excel样式是否已加载
        const existingLink = document.querySelector('link[href*="excel-styles.css"]');
        if (!existingLink) {
            const link = document.createElement('link');
            link.rel = 'stylesheet';
            link.href = 'css/excel-styles.css';
            document.head.appendChild(link);
            console.log('✅ Excel样式文件已加载');
        }
    }

    /**
     * 使用UTF-8编码读取CSV文件
     * @param {File} file - CSV文件对象
     * @returns {Promise<string>} UTF-8编码的文本内容
     */
    async readCSVWithUTF8(file) {
        try {
            // 首先尝试直接以UTF-8读取
            let text = await this.readFileAsText(file, 'utf-8');
            
            // 检查是否有乱码（简单检测）
            if (this.hasEncodingIssues(text)) {
                
                // 尝试以GBK/GB2312编码读取
                try {
                    text = await this.readFileAsText(file, 'gbk');
                } catch (gbkError) {
                    
                    // 回退到使用ArrayBuffer手动解码
                    const arrayBuffer = await file.arrayBuffer();
                    text = this.decodeArrayBufferAsUTF8(arrayBuffer);
                }
            }
            
            return text;
        } catch (error) {
            // 最后的回退方案：强制以UTF-8读取
            const arrayBuffer = await file.arrayBuffer();
            return this.decodeArrayBufferAsUTF8(arrayBuffer);
        }
    }

    /**
     * 将文件读取为指定编码的文本
     * @param {File} file - 文件对象
     * @param {string} encoding - 编码格式
     * @returns {Promise<string>} 文本内容
     */
    async readFileAsText(file, encoding = 'utf-8') {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = () => reject(reader.error);
            reader.readAsText(file, encoding);
        });
    }

    /**
     * 检测文本是否有编码问题
     * @param {string} text - 文本内容
     * @returns {boolean} 是否有编码问题
     */
    hasEncodingIssues(text) {
        // 检查是否包含常见的编码错误字符
        const encodingErrorPatterns = [
            /�/g,  // 替换字符
            /â©/g, /çº/g, /â/g,  // 常见的UTF-8误读为Latin-1的模式
            /\ufffd/g  // Unicode替换字符
        ];
        
        for (const pattern of encodingErrorPatterns) {
            if (pattern.test(text)) {
                return true;
            }
        }
        
        return false;
    }

    /**
     * 将ArrayBuffer解码为UTF-8文本
     * @param {ArrayBuffer} arrayBuffer - 数组缓冲区
     * @returns {string} UTF-8文本
     */
    decodeArrayBufferAsUTF8(arrayBuffer) {
        try {
            // 使用TextDecoder进行UTF-8解码
            const decoder = new TextDecoder('utf-8', { fatal: false });
            const text = decoder.decode(arrayBuffer);
            return text;
        } catch (error) {
            // 强制解码：忽略错误字符
            const decoder = new TextDecoder('utf-8', { fatal: false, ignoreBOM: true });
            return decoder.decode(arrayBuffer);
        }
    }

    /**
     * 将工作表转换为干净的JSON格式
     * @param {Object} worksheet - SheetJS工作表对象
     * @returns {Array} 清理后的JSON数组
     */
    convertSheetToCleanJSON(worksheet) {
        try {
            if (!worksheet || !worksheet['!ref']) {
                return [];
            }

            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const result = [];
            
            // 读取所有数据到二维数组
            const rawData = [];
            for (let row = range.s.r; row <= range.e.r; row++) {
                const rowData = [];
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                    const cell = worksheet[cellAddress];
                    rowData.push(cell ? (cell.v || '') : '');
                }
                rawData.push(rowData);
            }

            if (rawData.length === 0) return [];

            // 检测是否有有效的标题行
            const hasHeaders = this.hasValidHeaders(rawData[0]);
            
            // 转换数据行
            for (let i = 0; i < rawData.length; i++) {
                const row = rawData[i];
                
                // 跳过完全空白的行
                if (this.isEmptyRow(row)) continue;
                
                const rowObject = {};
                let hasData = false;
                
                // 计算实际的Excel行号（从1开始）
                const excelRowNum = i + 1;
                
                for (let j = 0; j < row.length; j++) {
                    const value = row[j];
                    if (value !== '' && value !== null && value !== undefined) {
                        // 生成Excel样式的单元格地址，如A1, B1, C2等
                        const columnLetter = this.columnIndexToLetter(j);
                        const cellKey = `${columnLetter}${excelRowNum}`;
                        
                        rowObject[cellKey] = value;
                        hasData = true;
                    }
                }
                
                // 只添加有数据的行
                if (hasData) {
                    // 添加行信息
                    rowObject._row = excelRowNum;
                    if (hasHeaders && i === 0) {
                        rowObject._type = 'header';
                    } else {
                        rowObject._type = 'data';
                    }
                    result.push(rowObject);
                }
            }

            return result;
            
        } catch (error) {
            console.warn('JSON转换失败，使用默认方式:', error);
            // 回退到默认方式
            return XLSX.utils.sheet_to_json(worksheet, { header: 1 }).slice(1);
        }
    }

    /**
     * 生成清洁的列标题
     * @param {Array} rawData - 原始数据数组
     * @param {number} columnCount - 列数
     * @returns {Array} 标题数组
     */
    generateCleanHeaders(rawData, columnCount) {
        const headers = [];
        
        // 检查第一行是否可以作为标题
        if (rawData.length > 0 && this.hasValidHeaders(rawData[0])) {
            // 使用第一行作为标题，但需要清理
            for (let i = 0; i < columnCount; i++) {
                let header = rawData[0][i] || '';
                header = this.cleanHeaderName(header, i);
                headers.push(header);
            }
        } else {
            // 生成默认标题 (A, B, C, ...)
            for (let i = 0; i < columnCount; i++) {
                headers.push(this.columnIndexToLetter(i));
            }
        }
        
        return headers;
    }

    /**
     * 检查第一行是否包含有效的标题
     * @param {Array} firstRow - 第一行数据
     * @returns {boolean} 是否为有效标题行
     */
    hasValidHeaders(firstRow) {
        if (!firstRow || firstRow.length === 0) return false;
        
        let validHeaders = 0;
        for (const cell of firstRow) {
            if (cell && typeof cell === 'string' && cell.trim() !== '') {
                // 排除纯数字、特别长的字符串等不适合做标题的内容
                if (!/^\d+$/.test(cell.trim()) && cell.length < 50) {
                    validHeaders++;
                }
            }
        }
        
        // 如果至少有一半的列有合理的标题，就认为是标题行
        return validHeaders >= Math.ceil(firstRow.length / 2);
    }

    /**
     * 清理标题名称
     * @param {string} header - 原始标题
     * @param {number} index - 列索引
     * @returns {string} 清理后的标题
     */
    cleanHeaderName(header, index) {
        if (!header || typeof header !== 'string') {
            return this.columnIndexToLetter(index);
        }
        
        header = header.trim();
        
        // 如果是纯数字或太长，使用列字母
        if (/^\d+$/.test(header) || header.length > 30) {
            return this.columnIndexToLetter(index);
        }
        
        // 移除特殊字符，保留中英文和数字
        header = header.replace(/[^\w\u4e00-\u9fff]/g, '_');
        
        // 确保不为空
        if (header === '' || header === '_') {
            return this.columnIndexToLetter(index);
        }
        
        return header;
    }

    /**
     * 将列索引转换为字母 (0->A, 1->B, ...)
     * @param {number} index - 列索引
     * @returns {string} 列字母
     */
    columnIndexToLetter(index) {
        let result = '';
        while (index >= 0) {
            result = String.fromCharCode(65 + (index % 26)) + result;
            index = Math.floor(index / 26) - 1;
        }
        return result;
    }

    /**
     * 检查行是否为空
     * @param {Array} row - 行数据
     * @returns {boolean} 是否为空行
     */
    isEmptyRow(row) {
        return !row || row.every(cell => 
            cell === '' || cell === null || cell === undefined
        );
    }

    /**
     * HTML转义
     * @param {string} text - 需要转义的文本
     * @returns {string} 转义后的文本
     */
    escapeHtml(text) {
        const map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, m => map[m]);
    }
}

// 导出类
if (typeof window !== 'undefined') {
    window.ExcelViewer = ExcelViewer;
}
