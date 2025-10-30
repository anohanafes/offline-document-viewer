/**
 * 文档处理工具类 - Vue版本
 * 从原始document-viewer.js迁移而来
 */

export class DocumentUtils {
  /**
   * 格式化文件大小显示
   * @param {number} bytes - 文件大小（字节）
   * @returns {string} 格式化后的文件大小字符串
   */
  static formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes'
    const k = 1024
    const sizes = ['Bytes', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
  }

  /**
   * 获取文件扩展名
   * @param {string} fileName - 文件名
   * @returns {string} 小写的文件扩展名
   */
  static getFileExtension(fileName) {
    return fileName.split('.').pop().toLowerCase()
  }

  /**
   * 检查是否为支持的文件格式
   * @param {string} extension - 文件扩展名
   * @returns {boolean} 是否支持该格式
   */
  static isSupportedFormat(extension) {
    const supportedFormats = ['pdf', 'docx', 'pptx', 'xlsx', 'xls', 'xlsm', 'xlsb', 'csv']
    return supportedFormats.includes(extension)
  }

  /**
   * 动态加载脚本文件
   * @param {string} src - 脚本文件路径
   * @returns {Promise} 加载完成的Promise
   */
  static loadScript(src) {
    return new Promise((resolve, reject) => {
      // 检查是否已经加载
      if (document.querySelector(`script[src="${src}"]`)) {
        resolve()
        return
      }

      const script = document.createElement('script')
      script.src = src
      script.onload = resolve
      script.onerror = reject
      document.head.appendChild(script)
    })
  }

  /**
   * 动态加载CSS文件
   * @param {string} href - CSS文件路径
   * @returns {Promise} 加载完成的Promise
   */
  static loadCSS(href) {
    return new Promise((resolve, reject) => {
      // 检查是否已经加载
      if (document.querySelector(`link[href="${href}"]`)) {
        resolve()
        return
      }

      const link = document.createElement('link')
      link.rel = 'stylesheet'
      link.href = href
      link.onload = resolve
      link.onerror = reject
      document.head.appendChild(link)
    })
  }

  /**
   * 读取文件为ArrayBuffer
   * @param {File} file - 文件对象
   * @returns {Promise<ArrayBuffer>} 文件内容
   */
  static readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = (e) => resolve(e.target.result)
      reader.onerror = reject
      reader.readAsArrayBuffer(file)
    })
  }

  /**
   * 从URL下载文件
   * @param {string} url - 文件URL
   * @returns {Promise<ArrayBuffer>} 文件内容
   */
  static async downloadFile(url) {
    try {
      console.log('正在下载文件:', url)
      
      // 尝试直接fetch
      let response
      try {
        response = await fetch(url, {
          mode: 'cors',
          credentials: 'omit',
          headers: {
            'Accept': 'application/octet-stream, */*'
          }
        })
      } catch (corsError) {
        console.warn('CORS请求失败，尝试no-cors模式:', corsError.message)
        // 如果CORS失败，尝试no-cors模式（但这会限制响应数据）
        response = await fetch(url, {
          mode: 'no-cors'
        })
      }
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status} ${response.statusText}`)
      }
      
      const arrayBuffer = await response.arrayBuffer()
      console.log('文件下载成功，大小:', arrayBuffer.byteLength, 'bytes')
      
      return arrayBuffer
    } catch (error) {
      console.error('下载文件详细错误:', error)
      throw new Error(`下载文件失败: ${error.message}`)
    }
  }

  /**
   * 从URL获取文件名
   * @param {string} url - 文件URL
   * @returns {string} 文件名
   */
  static getFileNameFromUrl(url) {
    try {
      const urlObj = new URL(url)
      const pathname = urlObj.pathname
      const fileName = pathname.split('/').pop()
      return decodeURIComponent(fileName || 'document')
    } catch (error) {
      return 'document'
    }
  }

  /**
   * 检测文件编码（主要用于CSV）
   * @param {ArrayBuffer} buffer - 文件内容
   * @returns {string} 编码类型
   */
  static detectEncoding(buffer) {
    const bytes = new Uint8Array(buffer)
    
    // 检测BOM
    if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
      return 'UTF-8'
    }
    if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
      return 'UTF-16LE'
    }
    if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
      return 'UTF-16BE'
    }

    // 简单的UTF-8检测
    let isUTF8 = true
    for (let i = 0; i < Math.min(bytes.length, 1000); i++) {
      if (bytes[i] > 127) {
        isUTF8 = false
        break
      }
    }

    return isUTF8 ? 'UTF-8' : 'GBK'
  }

  /**
   * 创建错误信息显示
   * @param {string} message - 错误消息
   * @param {string} details - 错误详情
   * @returns {string} HTML错误内容
   */
  static createErrorDisplay(message, details = '') {
    return `
      <div class="error-display">
        <div class="error-icon">❌</div>
        <h3 class="error-title">文档加载失败</h3>
        <p class="error-message">${message}</p>
        ${details ? `<div class="error-details">${details}</div>` : ''}
        <div class="error-actions">
          <button onclick="location.reload()" class="btn btn-primary">重新加载</button>
        </div>
      </div>
    `
  }

  /**
   * 节流函数
   * @param {Function} func - 要执行的函数
   * @param {number} delay - 延迟时间（毫秒）
   * @returns {Function} 节流后的函数
   */
  static throttle(func, delay) {
    let timeoutId
    let lastExecTime = 0
    return function (...args) {
      const currentTime = Date.now()
      
      if (currentTime - lastExecTime > delay) {
        func.apply(this, args)
        lastExecTime = currentTime
      } else {
        clearTimeout(timeoutId)
        timeoutId = setTimeout(() => {
          func.apply(this, args)
          lastExecTime = Date.now()
        }, delay - (currentTime - lastExecTime))
      }
    }
  }

  /**
   * 防抖函数
   * @param {Function} func - 要执行的函数
   * @param {number} delay - 延迟时间（毫秒）
   * @returns {Function} 防抖后的函数
   */
  static debounce(func, delay) {
    let timeoutId
    return function (...args) {
      clearTimeout(timeoutId)
      timeoutId = setTimeout(() => func.apply(this, args), delay)
    }
  }
}

export default DocumentUtils
