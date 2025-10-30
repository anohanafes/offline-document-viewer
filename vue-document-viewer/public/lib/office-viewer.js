/**
 * 基于第三方库的Office文档预览器
 * 使用更成熟的解决方案来处理PPTX预览
 */

class OfficeViewer {
    constructor() {
        this.currentFile = null;
        this.slides = [];
    }
    
    // 使用Office-Viewer风格的PPTX预览
    async renderPPTXWithOfficeViewer(file, container) {
        try {
            
             // 创建预览容器
             container.innerHTML = `
                 <div class="office-viewer-content" id="officeViewerContent">
                     <div class="loading-message">
                         <div class="loading-spinner"></div>
                         <p>正在解析PPTX文件...</p>
                     </div>
                 </div>
             `;
            
            // 尝试使用FileReader读取文件
            const arrayBuffer = await file.arrayBuffer();
            const blob = new Blob([arrayBuffer], { 
                type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' 
            });
            
            // 创建临时URL用于嵌入式预览
            const blobUrl = URL.createObjectURL(blob);
            
            // 尝试使用不同的预览方法
            await this.tryMultiplePreviewMethods(blobUrl, file, container);
            
        } catch (error) {
            console.error('Office-Viewer预览失败:', error);
            this.showFallbackPreview(container, file);
        }
    }
    
    async tryMultiplePreviewMethods(blobUrl, file, container) {
        const contentDiv = document.getElementById('officeViewerContent');
       
        // 使用组合预览：同时提取图片和文本
        try {
            await this.tryCombinedPreview(file, contentDiv);
            return;
        } catch (error) {
            console.warn('组合预览失败:', error);
        }
        
        // 备用方案1: 仅图片提取
        try {
            await this.tryImageExtractionPreview(file, contentDiv);
            return;
        } catch (error) {
            console.warn('图片提取预览失败:', error);
        }
        
        // 备用方案2: 仅文本提取
        try {
            await this.tryTextExtractionPreview(file, contentDiv);
            return;
        } catch (error) {
            console.warn('文本提取预览失败:', error);
        }
        
        // 最后的备用方案
        this.showFallbackPreview(container, file);
    }
    
    // 组合预览：同时显示图片和文本，保持原始布局
    async tryCombinedPreview(file, container) {
        
        
        const arrayBuffer = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);
        
        console.log('zip', zip);
        
        // 先提取所有图片
        const allImages = await this.extractImages(zip);
        
        // 解析幻灯片关系文件以建立图片映射
        const slideRelations = await this.extractSlideRelations(zip);
        
        // 提取幻灯片内容，包括图片和文本的位置关系
        const slides = await this.extractSlidesWithMedia(zip, allImages, slideRelations);
        // debugger;
        this.renderSlidesWithLayout(slides, container);
    }
    
    // 提取图片
    async extractImages(zip) {
        const images = [];
        const imagePromises = [];
        
        zip.folder('ppt/media/').forEach((relativePath, file) => {
            if (relativePath.match(/\.(jpg|jpeg|png|gif|bmp|svg|webp)$/i)) {
                const imagePromise = file.async('blob').then(imageData => {
                    const imageUrl = URL.createObjectURL(imageData);
                    return {
                        name: relativePath,
                        url: imageUrl,
                        size: imageData.size
                    };
                }).catch(error => {
                    console.warn('提取图片失败:', relativePath, error);
                    return null;
                });
                imagePromises.push(imagePromise);
            }
        });
        
        const imageResults = await Promise.all(imagePromises);
        return imageResults.filter(img => img !== null);
    }
    
    // 提取幻灯片关系文件
    async extractSlideRelations(zip) {
        const relations = {};
        const relPromises = [];
        
        zip.folder('ppt/slides/_rels/').forEach((relativePath, file) => {
            if (relativePath.endsWith('.xml.rels')) {
                const slideNumber = relativePath.match(/slide(\d+)\.xml\.rels/)?.[1];
                if (slideNumber) {
                    const relPromise = file.async('text').then(relsXml => {
                        relations[slideNumber] = this.parseRelationships(relsXml);
                    }).catch(error => {
                        console.warn('解析关系文件失败:', relativePath, error);
                    });
                    relPromises.push(relPromise);
                }
            }
        });
        
        // 等待所有关系文件解析完成
        await Promise.all(relPromises);
        
        return relations;
    }
    
    // 解析关系XML
    parseRelationships(xmlContent) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlContent, 'text/xml');
        const relationships = {};
        
        const relElements = xmlDoc.getElementsByTagName('Relationship');
        for (let i = 0; i < relElements.length; i++) {
            const rel = relElements[i];
            const id = rel.getAttribute('Id');
            const target = rel.getAttribute('Target');
            const type = rel.getAttribute('Type');
            
            if (type && (type.includes('image') || type.includes('Image') || target.match(/\.(jpg|jpeg|png|gif|svg|bmp|webp)$/i))) {
                relationships[id] = {
                    target: target,
                    type: 'image'
                };
            }
        }
        
        return relationships;
    }
    
    // 提取幻灯片内容，包括图片和文本的布局关系
    async extractSlidesWithMedia(zip, allImages, slideRelations) {
        const slides = [];
        const slideFiles = [];
        zip.folder('ppt/slides/').forEach((relativePath, file) => {
            if (relativePath.endsWith('.xml') && !relativePath.includes('_rels/')) {
                slideFiles.push({
                    path: relativePath,
                    file: file,
                    index: parseInt(relativePath.match(/slide(\d+)\.xml/)?.[1] || '0')
                });
            }
        });
        
        slideFiles.sort((a, b) => a.index - b.index);

        for (let i = 0; i < slideFiles.length; i++) {
            try {
                const slideXml = await slideFiles[i].file.async('text');
                const slideIndex = slideFiles[i].index;
                const slideContent = this.extractSlideContentWithLayout(
                    slideXml, 
                    slideIndex, 
                    allImages, 
                    slideRelations[slideIndex] || {}
                );
                slides.push(slideContent);
            } catch (error) {
                console.warn(`解析幻灯片${i + 1}失败:`, error);
                slides.push({
                    index: i + 1,
                    title: `幻灯片 ${i + 1}`,
                    elements: [{type: 'text', content: '解析失败'}]
                });
            }
        }
        
        return slides;
    }
    
    // 提取单张幻灯片的完整内容和布局
    extractSlideContentWithLayout(xmlContent, slideIndex, allImages, relations) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlContent, 'text/xml');
        
        
        const slide = {
            index: slideIndex,
            title: `幻灯片 ${slideIndex}`,
            elements: []
        };
        
        // 简化的元素提取逻辑
        const spShapes = xmlDoc.getElementsByTagName('p:sp');
        const picElements = xmlDoc.getElementsByTagName('p:pic');
        const allElements = [];
        
        // 提取文本形状
        for (let i = 0; i < spShapes.length; i++) {
            const shape = spShapes[i];
            
            // 使用智能文本提取方法，能分离多个文本块
            const textBlocks = this.extractShapeTextSmart(shape);
            
            textBlocks.forEach((textBlock, blockIndex) => {
                if (textBlock.content.trim()) {
                    allElements.push({
                        type: 'text',
                        content: textBlock.content,
                        formattedContent: textBlock.formattedContent,
                        position: textBlock.position,
                        order: this.getElementOrder(shape) + blockIndex * 0.1,
                        isTitle: this.isLikelyTitle(textBlock.content)
                    });
                }
            });
        }
        
        // 提取图片元素
        for (let i = 0; i < picElements.length; i++) {
            const pic = picElements[i];
            const imageInfo = this.extractImageInfo(pic, relations, allImages);
            if (imageInfo) {
                allElements.push({
                    type: 'image',
                    content: imageInfo,
                    position: imageInfo.position,
                    order: this.getElementOrder(pic)
                });
            }
        }
        
        allElements.sort((a, b) => a.order - b.order);
        
        // 设置幻灯片标题
        const titleElement = allElements.find(el => el.type === 'text' && el.isTitle) || 
                            allElements.find(el => el.type === 'text');
        slide.title = titleElement ? titleElement.content : `幻灯片 ${slideIndex}`;
        
        slide.elements = allElements;
        return slide;
    }
    
    // 提取形状中的文本段落和位置（支持独立文本框架）
    extractShapeTextParagraphs(shapeElement) {
        const results = [];
        
        // 查找所有文本框架 (p:txBody)
        const txBodyElements = shapeElement.getElementsByTagName('p:txBody');
        
        if (txBodyElements.length > 0) {
            // 有文本框架，按段落解析
            const txBody = txBodyElements[0];
            const paragraphs = txBody.getElementsByTagName('a:p');
            
            for (let i = 0; i < paragraphs.length; i++) {
                const paragraph = paragraphs[i];
                const textElements = paragraph.getElementsByTagName('a:t');
                const texts = [];
                
                for (let j = 0; j < textElements.length; j++) {
                    const text = textElements[j].textContent.trim();
                    if (text) {
                        texts.push(text);
                    }
                }
                
                if (texts.length > 0) {
                    const basePosition = this.extractElementPosition(shapeElement);
                    
                    // 为每个段落计算偏移位置
                    const paragraphPosition = {
                        ...basePosition,
                        y: basePosition.y + (i * 25), // 简单的垂直偏移
                        height: Math.max(25, basePosition.height / Math.max(paragraphs.length, 1))
                    };
                    
                    const content = texts.join(' ');
                    
                    results.push({
                        content: content,
                        position: paragraphPosition
                    });
                }
            }
        } else {
            // 没有文本框架，使用原有逻辑
            const textElements = shapeElement.getElementsByTagName('a:t');
            const texts = [];
            
            for (let i = 0; i < textElements.length; i++) {
                const text = textElements[i].textContent.trim();
                if (text) {
                    texts.push(text);
                }
            }
            
            if (texts.length > 0) {
                const position = this.extractElementPosition(shapeElement);
                const content = texts.join(' ');
                
                results.push({
                    content: content,
                    position: position
                });
            }
        }
        
        return results;
    }
    
    // 提取独立的文本运行（更细粒度的文本解析）
    extractIndividualTextRuns(shapeElement) {
        const results = [];
        const basePosition = this.extractElementPosition(shapeElement);
        
        // 查找所有文本运行 (a:r)
        const textRuns = shapeElement.getElementsByTagName('a:r');
        
        for (let i = 0; i < textRuns.length; i++) {
            const run = textRuns[i];
            const textElements = run.getElementsByTagName('a:t');
            const texts = [];
            
            for (let j = 0; j < textElements.length; j++) {
                const text = textElements[j].textContent.trim();
                if (text) {
                    texts.push(text);
                }
            }
            
            if (texts.length > 0) {
                const content = texts.join(' ');
                
                // 尝试从文本运行的属性中获取更精确的位置
                // 如果没有特定位置，使用基础位置并添加偏移
                const runPosition = {
                    ...basePosition,
                    x: basePosition.x + (i * Math.max(100, basePosition.width / Math.max(textRuns.length, 1))),
                    width: Math.max(80, basePosition.width / Math.max(textRuns.length, 1))
                };
                
                
                results.push({
                    content: content,
                    position: runPosition
                });
            }
        }
        
        return results;
    }
    
    // 智能文本提取方法 - 按段落分离独立文本块，并支持格式
    extractShapeTextSmart(shapeElement) {
        const results = [];
        const basePosition = this.extractElementPosition(shapeElement);
        
        // 首先尝试按段落(a:p)分离
        const paragraphs = shapeElement.getElementsByTagName('a:p');
        
        if (paragraphs.length > 1) {
            // 多个段落，分别处理，每个段落提取格式
            for (let i = 0; i < paragraphs.length; i++) {
                const para = paragraphs[i];
                
                // 为单个段落提取格式化内容
                const formattedContent = this.extractFormattedTextFromParagraph(para);
                
                if (formattedContent.text.trim()) {
                    // 尝试从段落本身获取位置信息
                    let paraPosition = this.extractElementPosition(para);
                    
                    // 如果段落没有独立位置信息，计算段落在文本框内的相对位置
                    if (paraPosition.x === 50 && paraPosition.y === 50) {
                        // 获取段落属性来计算垂直位置
                        const verticalOffset = this.calculateParagraphVerticalOffset(para, i, paragraphs.length, basePosition);
                        
                        paraPosition = {
                            ...basePosition,
                            y: basePosition.y + verticalOffset,
                            height: Math.max(25, basePosition.height / Math.max(paragraphs.length, 1))
                        };
                    }
                    
                    results.push({
                        content: formattedContent.text,
                        formattedContent: formattedContent,
                        position: paraPosition
                    });
                }
            }
        } 
        else {
            // 单个段落，尝试按文本运行(a:r)分离  
            const textRuns = shapeElement.getElementsByTagName('a:r');
            
            if (textRuns.length > 1) {
                // 多个文本运行，分别处理
                for (let i = 0; i < textRuns.length; i++) {
                    const run = textRuns[i];
                    
                    // 为单个文本运行提取格式化内容
                    const formattedContent = this.extractFormattedTextFromRun(run);
                    
                    if (formattedContent.text.trim()) {
                        // 尝试从文本运行本身获取位置信息
                        let runPosition = this.extractElementPosition(run);
                        
                        // 如果文本运行没有独立位置信息，直接使用形状的基础位置
                        if (runPosition.x === 50 && runPosition.y === 50) {
                            runPosition = basePosition;
                        }
                        
                        results.push({
                            content: formattedContent.text,
                            formattedContent: formattedContent,
                            position: runPosition
                        });
                    }
                }
            } else {
                // 单个文本块，使用完整的格式化提取
                const formattedContent = this.extractFormattedTextContent(shapeElement);
                
                if (formattedContent.text.trim()) {
                    results.push({
                        content: formattedContent.text,
                        formattedContent: formattedContent,
                        position: basePosition
                    });
                }
            }
        }
        
        return results;
    }
    
    // 从单个段落提取格式化文本
    extractFormattedTextFromParagraph(paragraph) {
        const result = {
            text: '',
            segments: []
        };
        
        // 获取段落样式
        const paragraphStyle = this.extractParagraphStyle(paragraph);
        
        // 获取文本运行
        const runs = paragraph.getElementsByTagName('a:r');
        
        if (runs.length === 0) {
            // 直接提取文本
            const textElements = paragraph.getElementsByTagName('a:t');
            for (let t = 0; t < textElements.length; t++) {
                const textContent = textElements[t].textContent;
                if (textContent) {
                    result.segments.push({
                        text: textContent,
                        style: { ...paragraphStyle }
                    });
                    result.text += textContent;
                }
            }
        } else {
            for (let r = 0; r < runs.length; r++) {
                const run = runs[r];
                const runStyle = this.extractRunStyle(run);
                
                const textElements = run.getElementsByTagName('a:t');
                for (let t = 0; t < textElements.length; t++) {
                    const textContent = textElements[t].textContent;
                    if (textContent) {
                        result.segments.push({
                            text: textContent,
                            style: { ...paragraphStyle, ...runStyle }
                        });
                        result.text += textContent;
                    }
                }
            }
        }
        
        return result;
    }
    
    // 从单个文本运行提取格式化文本
    extractFormattedTextFromRun(run) {
        const result = {
            text: '',
            segments: []
        };
        
        const runStyle = this.extractRunStyle(run);
        
        const textElements = run.getElementsByTagName('a:t');
        for (let t = 0; t < textElements.length; t++) {
            const textContent = textElements[t].textContent;
            if (textContent) {
                result.segments.push({
                    text: textContent,
                    style: { ...runStyle }
                });
                result.text += textContent;
            }
        }
        
        return result;
    }
    
    // 增强的文本提取方法 - 支持格式信息
    extractShapeTextSimple(shapeElement) {
        const position = this.extractElementPosition(shapeElement);
        
        // 提取带格式的文本内容
        const formattedContent = this.extractFormattedTextContent(shapeElement);
        
        return {
            content: formattedContent.text,
            formattedContent: formattedContent,
            position: position
        };
    }
    
    // 提取带格式的文本内容
    extractFormattedTextContent(shapeElement) {
        
        const result = {
            text: '',
            segments: [] // 存储每个文本段的内容和格式
        };
        
        // 获取所有段落
        const paragraphs = shapeElement.getElementsByTagName('a:p');
        
        for (let p = 0; p < paragraphs.length; p++) {
            const paragraph = paragraphs[p];
            
            // 获取段落级别的格式
            const paragraphStyle = this.extractParagraphStyle(paragraph);
            
            // 获取文本运行 (a:r)
            const runs = paragraph.getElementsByTagName('a:r');
            
            if (runs.length === 0) {
                // 如果没有文本运行，直接提取文本
                const textElements = paragraph.getElementsByTagName('a:t');
                for (let t = 0; t < textElements.length; t++) {
                    const textContent = textElements[t].textContent;
                    if (textContent) {
                        result.segments.push({
                            text: textContent,
                            style: { ...paragraphStyle }
                        });
                        result.text += textContent;
                    }
                }
            } else {
                for (let r = 0; r < runs.length; r++) {
                    const run = runs[r];
                    
                    // 提取文本运行的格式属性
                    const runStyle = this.extractRunStyle(run);
                    
                    // 获取文本内容
                    const textElements = run.getElementsByTagName('a:t');
                    
                    for (let t = 0; t < textElements.length; t++) {
                        const textContent = textElements[t].textContent;
                        if (textContent) {
                            const combinedStyle = { ...paragraphStyle, ...runStyle };
                            
                            result.segments.push({
                                text: textContent,
                                style: combinedStyle
                            });
                            result.text += textContent;
                        }
                    }
                }
            }
            
            // 段落间添加换行（除了最后一个段落）
            if (p < paragraphs.length - 1) {
                result.text += ' ';
                result.segments.push({
                    text: ' ',
                    style: { isLineBreak: true }
                });
            }
        }
        
        return result;
    }
    
    // 提取段落样式
    extractParagraphStyle(paragraph) {
        const style = {};
        
        try {
            const pPrElements = paragraph.getElementsByTagName('a:pPr');
            if (pPrElements.length > 0) {
                const pPr = pPrElements[0];
                
                // 对齐方式
                const alignment = pPr.getAttribute('algn');
                if (alignment) {
                    switch (alignment) {
                        case 'l': style.textAlign = 'left'; break;
                        case 'ctr': style.textAlign = 'center'; break;
                        case 'r': style.textAlign = 'right'; break;
                        case 'just': style.textAlign = 'justify'; break;
                    }
                }
                
                // 缩进
                const marL = pPr.getAttribute('marL');
                if (marL) {
                    style.marginLeft = Math.round(parseInt(marL) / 914400 * 96) + 'px';
                }
            }
        } catch (error) {
            console.warn('提取段落样式失败:', error);
        }
        
        return style;
    }
    
    // 提取文本运行样式
    extractRunStyle(run) {
        const style = {};
        
        try {
            const rPrElements = run.getElementsByTagName('a:rPr');
            if (rPrElements.length > 0) {
                const rPr = rPrElements[0];
                
                
                // 字体大小 (sz属性，单位是半点，需要除以100)
                const fontSize = rPr.getAttribute('sz');
                if (fontSize) {
                    style.fontSize = Math.round(parseInt(fontSize) / 100) + 'pt';
                }
                
                // 加粗
                const bold = rPr.getAttribute('b');
                if (bold === '1') {
                    style.fontWeight = 'bold';
                }
                
                // 斜体
                const italic = rPr.getAttribute('i');
                if (italic === '1') {
                    style.fontStyle = 'italic';
                }
                
                // 下划线
                const underline = rPr.getAttribute('u');
                if (underline && underline !== 'none') {
                    style.textDecoration = 'underline';
                }
                
                // 删除线
                const strike = rPr.getAttribute('strike');
                if (strike && strike !== 'noStrike') {
                    style.textDecoration = (style.textDecoration || '') + ' line-through';
                }
                
                // 字体名称
                const latinElements = rPr.getElementsByTagName('a:latin');
                if (latinElements.length > 0) {
                    const fontFamily = latinElements[0].getAttribute('typeface');
                    if (fontFamily) {
                        style.fontFamily = fontFamily;
                    }
                }
                
                // 中文字体
                const eaElements = rPr.getElementsByTagName('a:ea');
                if (eaElements.length > 0) {
                    const eaFontFamily = eaElements[0].getAttribute('typeface');
                    if (eaFontFamily) {
                        style.fontFamily = eaFontFamily; // 中文字体优先
                    }
                }
                
                // 字体颜色
                const solidFillElements = rPr.getElementsByTagName('a:solidFill');
                if (solidFillElements.length > 0) {
                    const colorElements = solidFillElements[0].getElementsByTagName('a:srgbClr');
                    if (colorElements.length > 0) {
                        const colorValue = colorElements[0].getAttribute('val');
                        if (colorValue) {
                            style.color = '#' + colorValue;
                        }
                    }
                    
                    // 也检查主题颜色
                    const schemeClrElements = solidFillElements[0].getElementsByTagName('a:schemeClr');
                    if (schemeClrElements.length > 0) {
                        const schemeVal = schemeClrElements[0].getAttribute('val');
                        // 简单的主题颜色映射
                        const themeColors = {
                            'dk1': '#000000',
                            'lt1': '#ffffff', 
                            'dk2': '#1F497D',
                            'lt2': '#EEECE1',
                            'accent1': '#4F81BD',
                            'accent2': '#F79646',
                            'accent3': '#9BBB59',
                            'accent4': '#8064A2',
                            'accent5': '#4BACC6',
                            'accent6': '#F24992'
                        };
                        if (themeColors[schemeVal]) {
                            style.color = themeColors[schemeVal];
                        }
                    }
                }
                
                // 文字背景色/高亮色 (a:highlight)
                const highlightElements = rPr.getElementsByTagName('a:highlight');
                if (highlightElements.length > 0) {
                    const highlight = highlightElements[0];
                    
                    // 检查RGB高亮色
                    const hlColorElements = highlight.getElementsByTagName('a:srgbClr');
                    if (hlColorElements.length > 0) {
                        const hlColorValue = hlColorElements[0].getAttribute('val');
                        if (hlColorValue) {
                            style.backgroundColor = '#' + hlColorValue;
                        }
                    }
                    
                    // 检查主题高亮色
                    const hlSchemeClrElements = highlight.getElementsByTagName('a:schemeClr');
                    if (hlSchemeClrElements.length > 0) {
                        const hlSchemeVal = hlSchemeClrElements[0].getAttribute('val');
                        const themeColors = {
                            'dk1': '#000000',
                            'lt1': '#ffffff', 
                            'dk2': '#1F497D',
                            'lt2': '#EEECE1',
                            'accent1': '#4F81BD',
                            'accent2': '#F79646',
                            'accent3': '#9BBB59',
                            'accent4': '#8064A2',
                            'accent5': '#4BACC6',
                            'accent6': '#F24992',
                            'hlink': '#0563C1',
                            'folHlink': '#954F72'
                        };
                        if (themeColors[hlSchemeVal]) {
                            style.backgroundColor = themeColors[hlSchemeVal];
                        }
                    }
                    
                    // 检查预设高亮色
                    const hlPrstClrElements = highlight.getElementsByTagName('a:prstClr');
                    if (hlPrstClrElements.length > 0) {
                        const hlPrstVal = hlPrstClrElements[0].getAttribute('val');
                        // 常见的预设高亮颜色
                        const presetColors = {
                            'yellow': '#FFFF00',
                            'lime': '#00FF00',
                            'cyan': '#00FFFF',
                            'magenta': '#FF00FF',
                            'blue': '#0000FF',
                            'red': '#FF0000',
                            'darkBlue': '#000080',
                            'darkCyan': '#008080',
                            'darkGreen': '#008000',
                            'darkMagenta': '#800080',
                            'darkRed': '#800000',
                            'darkYellow': '#808000',
                            'darkGray': '#808080',
                            'lightGray': '#C0C0C0',
                            'black': '#000000'
                        };
                        if (presetColors[hlPrstVal]) {
                            style.backgroundColor = presetColors[hlPrstVal];
                        }
                    }
                }
                
                // 检查是否有默认文本属性
                const defRPrElements = run.parentNode.getElementsByTagName('a:defRPr');
                if (defRPrElements.length > 0 && Object.keys(style).length === 0) {
                    const defRPr = defRPrElements[0];
                    
                    const defFontSize = defRPr.getAttribute('sz');
                    if (defFontSize) {
                        style.fontSize = Math.round(parseInt(defFontSize) / 100) + 'pt';
                    }
                    
                    // 检查默认字体
                    const defLatinElements = defRPr.getElementsByTagName('a:latin');
                    if (defLatinElements.length > 0) {
                        const defFontFamily = defLatinElements[0].getAttribute('typeface');
                        if (defFontFamily) {
                            style.fontFamily = defFontFamily;
                        }
                    }
                }
            } else {
            }
        } catch (error) {
            console.warn('提取文本运行样式失败:', error);
        }
        
        return style;
    }
    
    // 计算段落在文本框内的垂直偏移
    calculateParagraphVerticalOffset(paragraph, paragraphIndex, totalParagraphs, basePosition) {
        let verticalOffset = 0;
        
        try {
            // 策略1：根据段落属性计算间距
            const pPrElements = paragraph.getElementsByTagName('a:pPr');
            if (pPrElements.length > 0) {
                const pPr = pPrElements[0];
                
                // 检查段前间距 (a:spcBef)
                const spcBefElements = pPr.getElementsByTagName('a:spcBef');
                if (spcBefElements.length > 0) {
                    const spcPts = spcBefElements[0].getElementsByTagName('a:spcPts');
                    if (spcPts.length > 0) {
                        const beforeSpacing = parseInt(spcPts[0].getAttribute('val') || 0);
                        verticalOffset += Math.round(beforeSpacing / 100);
                    }
                }
                
                // 检查行距 (a:lnSpc)
                const lnSpcElements = pPr.getElementsByTagName('a:lnSpc');
                if (lnSpcElements.length > 0) {
                    const spcPts = lnSpcElements[0].getElementsByTagName('a:spcPts');
                    if (spcPts.length > 0) {
                        const lineSpacing = parseInt(spcPts[0].getAttribute('val') || 0);
                        if (paragraphIndex > 0) {
                            verticalOffset += Math.round(lineSpacing / 100 * paragraphIndex);
                        }
                    }
                }
            }
            
            // 策略2：如果没有具体间距信息，使用默认计算
            if (verticalOffset === 0 && totalParagraphs > 1) {
                // 假设每行约25-30px的高度
                const estimatedLineHeight = 30;
                verticalOffset = paragraphIndex * estimatedLineHeight;
            }
            
        } catch (error) {
            console.warn('计算段落垂直偏移失败:', error);
            // 回退到简单的等分计算
            if (totalParagraphs > 1) {
                verticalOffset = (basePosition.height / totalParagraphs) * paragraphIndex;
            }
        }
        
        return verticalOffset;
    }
    
    // 生成格式化的HTML文本
    generateFormattedTextHtml(element) {
        if (!element.formattedContent || !element.formattedContent.segments) {
            return this.escapeHtml(element.content);
        }
        
        let html = '';
        element.formattedContent.segments.forEach((segment, index) => {
            if (segment.style.isLineBreak) {
                html += '<br>';
                return;
            }
            
            const styles = [];
            const style = segment.style;
            
            
            // 应用各种样式
            if (style.fontSize) styles.push(`font-size: ${style.fontSize}`);
            if (style.fontFamily) styles.push(`font-family: "${style.fontFamily}"`);
            if (style.fontWeight) styles.push(`font-weight: ${style.fontWeight}`);
            if (style.fontStyle) styles.push(`font-style: ${style.fontStyle}`);
            if (style.color) styles.push(`color: ${style.color}`);
            if (style.backgroundColor) styles.push(`background-color: ${style.backgroundColor}`);
            if (style.textDecoration) styles.push(`text-decoration: ${style.textDecoration}`);
            if (style.textAlign) styles.push(`text-align: ${style.textAlign}`);
            if (style.marginLeft) styles.push(`margin-left: ${style.marginLeft}`);
            
            const styleAttr = styles.length > 0 ? ` style="${styles.join('; ')}"` : '';
            
            html += `<span${styleAttr}>${this.escapeHtml(segment.text)}</span>`;
        });
        
        return html;
    }
    
    // 保持原有方法兼容性
    extractShapeText(shapeElement) {
        return this.extractShapeTextSimple(shapeElement);
    }
    
    // 提取图片信息和位置
    extractImageInfo(picElement, relations, allImages) {
        // 获取图片引用ID
        const blipElements = picElement.getElementsByTagName('a:blip');
        if (blipElements.length === 0) return null;
        
        const imageRefId = blipElements[0].getAttribute('r:embed');
        if (!imageRefId || !relations[imageRefId]) return null;
        
        const imagePath = relations[imageRefId].target;
        
        // 在allImages中查找对应的图片
        const matchingImage = allImages.find(img => 
            img.name.includes(imagePath.split('/').pop()) || 
            imagePath.includes(img.name.split('/').pop())
        );
        
        if (matchingImage) {
            // 提取位置和尺寸信息
            const position = this.extractElementPosition(picElement);
            
            return {
                url: matchingImage.url,
                name: matchingImage.name,
                size: matchingImage.size,
                alt: this.extractImageAlt(picElement),
                position: position
            };
        }
        
        return null;
    }
    
    // 提取图片替代文本
    extractImageAlt(picElement) {
        const descElements = picElement.getElementsByTagName('p:cNvPr');
        if (descElements.length > 0) {
            return descElements[0].getAttribute('descr') || 
                   descElements[0].getAttribute('name') || '';
        }
        return '';
    }
    
    // 提取元素位置信息 - 增强版，支持多种位置信息来源
    extractElementPosition(element) {
        try {
            // 策略1：查找 p:spPr > a:xfrm (形状属性中的变换)
            let xfrmElement = null;
            const spPrElements = element.getElementsByTagName('p:spPr');
            if (spPrElements.length > 0) {
                const xfrmInSpPr = spPrElements[0].getElementsByTagName('a:xfrm');
                if (xfrmInSpPr.length > 0) {
                    xfrmElement = xfrmInSpPr[0];
                }
            }
            
            // 策略2：查找 p:spPr > a:prstGeom > a:xfrm
            if (!xfrmElement && spPrElements.length > 0) {
                const prstGeomElements = spPrElements[0].getElementsByTagName('a:prstGeom');
                if (prstGeomElements.length > 0) {
                    const xfrmInGeom = prstGeomElements[0].getElementsByTagName('a:xfrm');
                    if (xfrmInGeom.length > 0) {
                        xfrmElement = xfrmInGeom[0];
                    }
                }
            }
            
            // 策略3：直接查找 a:xfrm
            if (!xfrmElement) {
                const xfrmElements = element.getElementsByTagName('a:xfrm');
                if (xfrmElements.length > 0) {
                    xfrmElement = xfrmElements[0];
                }
            }
            
            if (!xfrmElement) {
                return { x: 50, y: 50, width: 200, height: 50 };
            }
            
            // 提取偏移量 (a:off)
            const offElements = xfrmElement.getElementsByTagName('a:off');
            const x = offElements.length > 0 ? parseInt(offElements[0].getAttribute('x') || 0) : 0;
            const y = offElements.length > 0 ? parseInt(offElements[0].getAttribute('y') || 0) : 0;
            
            // 提取尺寸 (a:ext)
            const extElements = xfrmElement.getElementsByTagName('a:ext');
            const width = extElements.length > 0 ? parseInt(extElements[0].getAttribute('cx') || 0) : 0;
            const height = extElements.length > 0 ? parseInt(extElements[0].getAttribute('cy') || 0) : 0;
            
            // 将EMU转换为像素 (1 EMU = 1/914400 英寸, 1英寸 = 96像素)
            const emuToPixels = (emu) => Math.round(emu / 914400 * 96);
            
            const result = {
                x: emuToPixels(x),
                y: emuToPixels(y),
                width: emuToPixels(width) || 200,
                height: emuToPixels(height) || 50
            };
            
            
            return result;
        } catch (error) {
            console.error('提取位置信息失败:', error);
            return { x: 50, y: 50, width: 200, height: 50 };
        }
    }
    
    // 获取元素在XML中的顺序（简化版）
    getElementOrder(element) {
        let order = 0;
        let current = element;
        while (current.previousElementSibling) {
            order++;
            current = current.previousElementSibling;
        }
        return order;
    }
    
    // 判断是否可能是标题
    isLikelyTitle(text) {
        if (!text) return false;
        return text.length < 100 && 
               text.length > 2 && 
               !text.includes('。') && 
               !text.includes('!')&& 
               !text.includes('？');
    }
    
    // 渲染幻灯片，PowerPoint风格布局
    renderSlidesWithLayout(slides, container) {
        const totalImages = slides.reduce((count, slide) => {
            return count + slide.elements.filter(el => el.type === 'image').length;
        }, 0);
        
        this.slides = slides; // 保存slides数据供后续使用
        this.currentSlideIndex = 0; // 当前选中的幻灯片
        
        // 计算总页数
        const totalSlides = slides.length;
        
         let html = `
             <div class="ppt-layout-container">
                 <!-- 头部工具栏 -->
                 <div class="ppt-header">
                     <div class="ppt-toolbar">
                         <div class="zoom-controls">
                             <button class="tool-button zoom-out" id="zoomOutButton">🔍-</button>
                             <span class="zoom-level" id="zoomLevel">100%</span>
                             <button class="tool-button zoom-in" id="zoomInButton">🔍+</button>
                             <button class="tool-button fit-width" id="fitWidthButton">适应宽度</button>
                         </div>
                         <button class="tool-button fullscreen" id="fullscreenButton">全屏</button>
                     </div>
                 </div>
                
                <!-- 主要内容区域 -->
                <div class="ppt-main-content">
                    <!-- 左侧缩略图导航 - 已隐藏 -->
                    <div class="ppt-sidebar" style="display: none;">
                        <div class="sidebar-header">幻灯片</div>
                        <div class="thumbnails-container">
        `;
        
        // 生成缩略图列表
        slides.forEach((slide, index) => {

            const isActive = index === 0 ? 'active' : '';
            const hasImage = slide.elements.some(el => el.type === 'image');
            
            html += `
                <div class="thumbnail-slide ${isActive}" data-slide-index="${index}">
                    <div class="thumbnail-number">${index + 1}</div>
                    <div class="thumbnail-content">
                        <div class="thumbnail-title">${this.escapeHtml(slide.title)}</div>
                        <div class="thumbnail-preview">
            `;
            
            // 显示缩略图内容预览
            if (hasImage) {
                const firstImage = slide.elements.find(el => el.type === 'image');
                if (firstImage) {
                    html += `<img src="${firstImage.content.url}" class="thumbnail-image" alt="预览">`;
                }
            } else {
                html += `<div class="thumbnail-text">📝</div>`;
            }
            
            html += `
                        </div>
                    </div>
                </div>
            `;
        });
        
        html += `
                        </div>
                    </div>
                    
                     <!-- 右侧主预览区域 -->
                     <div class="ppt-preview-area" style="width: 100%; flex: 1;">
                         <div class="slide-display-area" id="slideDisplayArea">
        `;
        
        // 显示第一张幻灯片的内容
        html += this.renderSlideContent(slides[0]);
        
        html += `
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        container.innerHTML = html;
        
        // 绑定缩略图点击事件和工具栏按钮
        this.bindThumbnailEvents();
        
        // 页码显示已移除
         this.bindNavigationButtons();
        
    }
    
    // 渲染单张幻灯片内容 - 使用绝对定位保持PPT布局
    renderSlideContent(slide) {
        if (slide.elements.length === 0) {
            return `<div class="slide-content-display"><div class="display-empty">此幻灯片无可显示内容</div></div>`;
        }
        
        // 计算幻灯片边界
        const bounds = this.calculateSlideBounds(slide.elements);
        
        let html = `
            <div class="slide-content-display">
                <div class="slide-navigation">
                    <button class="nav-button prev-slide" id="prevSlideBtn">← 上一页</button>
                    <span class="slide-counter" id="slideCounter">第 ${slide.index} 页 / 共 ${this.slides.length} 页</span>
                    <button class="nav-button next-slide" id="nextSlideBtn">下一页 →</button>
                </div>
                <div class="slide-canvas" style="
                    position: relative;
                    width: ${bounds.width}px;
                    height: ${bounds.height}px;
                    margin: 0 auto;
                    background: #ffffff;
                    border: 1px solid #e1dfdd;
                    border-radius: 8px;
                    overflow: hidden;
                ">
        `;
        
        // 按类型分离并渲染元素（图片在底层，文字在上层）
        const images = slide.elements.filter(el => el.type === 'image');
        const texts = slide.elements.filter(el => el.type === 'text');
        
        // 先渲染图片元素（底层）
        images.forEach((element, imageIndex) => {
            const pos = element.position;
            const imageLeft = pos.x - bounds.minX;
            const imageTop = pos.y - bounds.minY;
            
            html += `
                <div class="positioned-image-container" style="
                    position: absolute;
                    left: ${pos.x - bounds.minX}px;
                    top: ${pos.y - bounds.minY}px;
                    width: ${pos.width}px;
                    height: ${pos.height}px;
                    z-index: 1;
                ">
                    <img src="${element.content.url}" 
                         alt="${this.escapeHtml(element.content.alt)}" 
                         style="
                            width: 100%;
                            height: 100%;
                            object-fit: cover;
                            border-radius: 4px;
                         ">
                </div>
            `;
        });
        
                // 再渲染文字元素（上层） - 支持格式化文本
        texts.forEach((element, textIndex) => {
            const pos = element.position;
            
            // 跳过已经作为标题显示的文本
            if (element.content !== slide.title) {
                const cssClass = element.isTitle ? 'positioned-subtitle' : 'positioned-text';
                
                // 使用原始PPT位置，不进行重叠调整
                let adjustedX = pos.x - bounds.minX;
                let adjustedY = pos.y - bounds.minY;
                
                
                
                // 生成格式化的HTML内容
                const formattedHtml = this.generateFormattedTextHtml(element);
                
                html += `
                    <div class="${cssClass}" style="
                        position: absolute;
                        left: ${adjustedX}px;
                        top: ${adjustedY}px;
                        width: auto;
                        height: auto;
                        z-index: 10;
                    ">${formattedHtml}</div>
                `;
            }
        });
        
        html += `
                </div>
            </div>
        `;
        
        return html;
    }
    
    // 计算幻灯片内容边界
    calculateSlideBounds(elements) {
        if (elements.length === 0) {
            return { minX: 0, minY: 0, maxX: 800, maxY: 600, width: 800, height: 600 };
        }
        
        let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        
        elements.forEach(element => {
            const pos = element.position;
            minX = Math.min(minX, pos.x);
            minY = Math.min(minY, pos.y);
            maxX = Math.max(maxX, pos.x + pos.width);
            maxY = Math.max(maxY, pos.y + pos.height);
        });
        
        // 添加边距
        const padding = 20;
        minX -= padding;
        minY -= padding;
        maxX += padding;
        maxY += padding;
        
        // 确保最小尺寸
        const width = Math.max(maxX - minX, 600);
        const height = Math.max(maxY - minY, 400);
        
        return { minX, minY, maxX, maxY, width, height };
    }
    
    // 绑定缩略图点击事件
    bindThumbnailEvents() {
        const thumbnails = document.querySelectorAll('.thumbnail-slide');
        const slideDisplayArea = document.getElementById('slideDisplayArea');
        const slideCounter = document.getElementById('slideCounter');
        
        // 初始化页码显示
        if (slideCounter) {
            slideCounter.textContent = `第 1 页 / 共 ${this.slides.length} 页`;
        }
        
        thumbnails.forEach(thumbnail => {
            thumbnail.addEventListener('click', () => {
                const slideIndex = parseInt(thumbnail.getAttribute('data-slide-index'));
                
                // 更新活跃状态
                thumbnails.forEach(t => t.classList.remove('active'));
                thumbnail.classList.add('active');
                
                // 更新右侧显示内容
                const selectedSlide = this.slides[slideIndex];
                slideDisplayArea.innerHTML = this.renderSlideContent(selectedSlide);
                
                // 更新页码显示
                if (slideCounter) {
                }
                
                this.currentSlideIndex = slideIndex;
            });
        });
        
        // 添加键盘导航支持
        document.addEventListener('keydown', (e) => {
            if ((e.key === 'ArrowUp' || e.key === 'ArrowLeft') && this.currentSlideIndex > 0) {
                thumbnails[this.currentSlideIndex - 1].click();
            } else if ((e.key === 'ArrowDown' || e.key === 'ArrowRight') && this.currentSlideIndex < this.slides.length - 1) {
                thumbnails[this.currentSlideIndex + 1].click();
            }
        });
    }
    
    // 绑定缩放和工具栏按钮事件
    bindNavigationButtons() {
        const thumbnails = document.querySelectorAll('.thumbnail-slide');
        const slideDisplayArea = document.getElementById('slideDisplayArea');
        const zoomInButton = document.getElementById('zoomInButton');
        const zoomOutButton = document.getElementById('zoomOutButton');
        const fitWidthButton = document.getElementById('fitWidthButton');
        const fullscreenButton = document.getElementById('fullscreenButton');
        const zoomLevelDisplay = document.getElementById('zoomLevel');
        
        // 绑定上一页/下一页按钮事件
        const bindSlideNavButtons = () => {
            const prevButton = document.getElementById('prevSlideBtn');
            const nextButton = document.getElementById('nextSlideBtn');
            
            if (prevButton) {
                prevButton.addEventListener('click', () => {
                    if (this.currentSlideIndex > 0) {
                        const prevIndex = this.currentSlideIndex - 1;
                        thumbnails[prevIndex].click();
                    }
                });
                prevButton.disabled = this.currentSlideIndex <= 0;
            }
            
            if (nextButton) {
                nextButton.addEventListener('click', () => {
                    if (this.currentSlideIndex < this.slides.length - 1) {
                        const nextIndex = this.currentSlideIndex + 1;
                        thumbnails[nextIndex].click();
                    }
                });
                nextButton.disabled = this.currentSlideIndex >= this.slides.length - 1;
            }
        };
        
        // 初始绑定导航按钮
        bindSlideNavButtons();
        
        // 每次更新幻灯片内容后重新绑定导航按钮
        const observer = new MutationObserver(() => {
            bindSlideNavButtons();
        });
        
        observer.observe(slideDisplayArea, { childList: true });
        
        // 初始化缩放级别
        let zoomLevel = 100;
        let slideCanvas = null;
        
        // 更新缩放级别显示
        const updateZoomDisplay = () => {
            zoomLevelDisplay.textContent = `${zoomLevel}%`;
            slideCanvas = document.querySelector('.slide-canvas');
            if (slideCanvas) {
                slideCanvas.style.transform = `scale(${zoomLevel / 100})`;
                slideCanvas.style.transformOrigin = 'center top';
            }
        };
        
        // 缩放按钮事件
        if (zoomInButton) {
            zoomInButton.addEventListener('click', () => {
                zoomLevel = Math.min(zoomLevel + 10, 200);
                updateZoomDisplay();
            });
        }
        
        if (zoomOutButton) {
            zoomOutButton.addEventListener('click', () => {
                zoomLevel = Math.max(zoomLevel - 10, 50);
                updateZoomDisplay();
            });
        }
        
        // 适应宽度按钮
        if (fitWidthButton) {
            fitWidthButton.addEventListener('click', () => {
                zoomLevel = 100;
                updateZoomDisplay();
                
                if (slideCanvas) {
                    const containerWidth = slideDisplayArea.clientWidth;
                    const canvasWidth = slideCanvas.clientWidth;
                    if (canvasWidth > 0) {
                        const fitRatio = (containerWidth - 40) / canvasWidth * 100;
                        zoomLevel = Math.min(Math.max(Math.round(fitRatio), 50), 150);
                        updateZoomDisplay();
                    }
                }
            });
        }
        
        // 全屏按钮
        if (fullscreenButton) {
            fullscreenButton.addEventListener('click', () => {
                const container = document.querySelector('.ppt-layout-container');
                if (container) {
                    if (!document.fullscreenElement) {
                        container.requestFullscreen().catch(err => {
                            console.error(`全屏错误: ${err.message}`);
                        });
                    } else {
                        document.exitFullscreen();
                    }
                }
            });
        }
    }
    
    // 设置标签页切换
    setupTabSwitching() {
        window.switchPreviewTab = (tabName) => {
            // 隐藏所有标签内容
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            
            // 移除所有按钮的active状态
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            
            // 显示选中的标签内容
            const targetTab = document.getElementById(tabName + 'Tab');
            if (targetTab) {
                targetTab.classList.add('active');
            }
            
            // 激活对应的按钮
            const buttons = document.querySelectorAll('.tab-btn');
            if (tabName === 'slides' && buttons[0]) {
                buttons[0].classList.add('active');
            } else if (tabName === 'images' && buttons[1]) {
                buttons[1].classList.add('active');
            }
        };
    }
    
    // 文本内容提取预览
    async tryTextExtractionPreview(file, container) {
        
        const arrayBuffer = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);
        const slides = [];
        
        // 获取幻灯片文件
        const slideFiles = [];
        zip.folder('ppt/slides/').forEach((relativePath, file) => {
            if (relativePath.endsWith('.xml') && !relativePath.includes('_rels/')) {
                slideFiles.push({
                    path: relativePath,
                    file: file,
                    index: parseInt(relativePath.match(/slide(\d+)\.xml/)?.[1] || '0')
                });
            }
        });
        
        // 按索引排序
        slideFiles.sort((a, b) => a.index - b.index);
        
        // 解析每个幻灯片的文本
        for (let i = 0; i < slideFiles.length; i++) {
            try {
                const slideXml = await slideFiles[i].file.async('text');
                const slideContent = this.extractSlideText(slideXml, i + 1);
                slides.push(slideContent);
            } catch (error) {
                console.warn(`解析幻灯片${i + 1}文本失败:`, error);
                slides.push({
                    index: i + 1,
                    title: `幻灯片 ${i + 1}`,
                    texts: ['无法解析内容']
                });
            }
        }
        
        // 渲染文本预览
        this.renderTextPreview(slides, container);
    }
    
    extractSlideText(xmlContent, slideIndex) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlContent, 'text/xml');
        
        const slide = {
            index: slideIndex,
            title: `幻灯片 ${slideIndex}`,
            texts: []
        };
        
        // 提取所有文本元素
        const textElements = xmlDoc.getElementsByTagName('a:t');
        const extractedTexts = [];
        
        for (let i = 0; i < textElements.length; i++) {
            const text = textElements[i].textContent.trim();
            if (text && text.length > 0) {
                extractedTexts.push(text);
            }
        }
        
        // 第一个文本通常是标题
        if (extractedTexts.length > 0) {
            slide.title = extractedTexts[0];
            slide.texts = extractedTexts.slice(1);
        }
        
        return slide;
    }
    
    renderTextPreview(slides, container) {
        let html = `
            <div class="text-extraction-preview">
                <div class="preview-notice">
                    <p>📝 PPTX文本内容预览</p>
                    <p>共解析 ${slides.length} 张幻灯片</p>
                </div>
                <div class="slides-text-content">
        `;
        
        slides.forEach(slide => {
            html += `
                <div class="slide-text-item">
                    <div class="slide-text-header">
                        <h3>${slide.title}</h3>
                        <span class="slide-number">第 ${slide.index} 页</span>
                    </div>
                    <div class="slide-text-body">
            `;
            
            if (slide.texts.length > 0) {
                slide.texts.forEach(text => {
                    html += `<p class="slide-text-paragraph">${this.escapeHtml(text)}</p>`;
                });
            } else {
                html += `<p class="slide-text-empty">此幻灯片无文本内容</p>`;
            }
            
            html += `
                    </div>
                </div>
            `;
        });
        
        html += `
                </div>
            </div>
        `;
        
        container.innerHTML = html;
    }
    
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    async tryImageExtractionPreview(file, container) {
        
        const arrayBuffer = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);
        
        // 提取所有图片
        const images = [];
        const imagePromises = [];
        
        // 获取媒体文件夹中的图片
        zip.folder('ppt/media/').forEach((relativePath, file) => {
            if (relativePath.match(/\.(jpg|jpeg|png|gif|bmp|svg|webp)$/i)) {
                const imagePromise = file.async('blob').then(imageData => {
                    const imageUrl = URL.createObjectURL(imageData);
                    return {
                        name: relativePath,
                        url: imageUrl,
                        size: imageData.size
                    };
                }).catch(error => {
                    console.warn('提取图片失败:', relativePath, error);
                    return null;
                });
                imagePromises.push(imagePromise);
            }
        });
        
        // 等待所有图片提取完成
        const imageResults = await Promise.all(imagePromises);
        images.push(...imageResults.filter(img => img !== null));
        
        if (images.length > 0) {
            let html = `
                <div class="image-extraction-preview">
                    <div class="preview-notice">
                        <p>🖼️ 从PPTX中提取的图片内容</p>
                        <p>共找到 ${images.length} 张图片</p>
                    </div>
                    <div class="extracted-images">
            `;
            
            images.forEach((image, index) => {
                html += `
                    <div class="extracted-image-item">
                        <div class="image-header">图片 ${index + 1}</div>
                        <img src="${image.url}" alt="${image.name}" style="max-width: 100%; height: auto; border-radius: 4px;">
                        <div class="image-caption">${image.name}</div>
                    </div>
                `;
            });
            
            html += `
                    </div>
                </div>
            `;
            
            container.innerHTML = html;
        } else {
            throw new Error('未找到可提取的图片');
        }
    }
    
    showFallbackPreview(container, file) {
        container.innerHTML = `
            <div class="fallback-preview">
                <div class="fallback-header">
                    <h3>📊 PowerPoint 文档信息</h3>
                </div>
                <div class="fallback-content">
                    <div class="file-info">
                        <p><strong>文件名:</strong> ${file.name}</p>
                        <p><strong>文件大小:</strong> ${this.formatFileSize(file.size)}</p>
                        <p><strong>文件类型:</strong> PPTX演示文稿</p>
                    </div>
                    
                    <div class="preview-options">
                        <h4>💡 建议的预览方案:</h4>
                        <ul>
                            <li>🔄 <strong>转换为PDF</strong> - 使用本地软件将PPTX转为PDF后预览</li>
                            <li>💻 <strong>本地软件</strong> - 使用PowerPoint、LibreOffice或WPS打开</li>
                            <li>📱 <strong>移动应用</strong> - 使用手机上的Office应用查看</li>
                            <li>🖼️ <strong>图片提取</strong> - 当前已显示文档中包含的图片内容</li>
                        </ul>
                    </div>
                    
                    <div class="technical-note">
                        <h4>🔧 技术说明:</h4>
                        <p>PPTX格式包含复杂的布局、动画和多媒体元素，完整的离线预览需要专门的渲染引擎。当前版本提供基础预览功能，未来会继续改进。</p>
                    </div>
                </div>
            </div>
        `;
    }
    
    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
}

// 全局导出
window.OfficeViewer = OfficeViewer;
