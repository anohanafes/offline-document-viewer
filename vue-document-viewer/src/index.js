import DocumentViewer from './components/DocumentViewer.vue'
import FileUploader from './components/FileUploader.vue'
import UrlInput from './components/UrlInput.vue'
import ViewerControls from './components/ViewerControls.vue'

// 导入样式
import './assets/styles/main.css'
import './assets/styles/excel-styles.css'

const components = {
  DocumentViewer,
  FileUploader,
  UrlInput,
  ViewerControls
}

// 自动安装插件
const install = function (Vue) {
  if (install.installed) return
  install.installed = true
  
  Object.keys(components).forEach(name => {
    Vue.component(name, components[name])
  })
}

// 自动安装（如果在浏览器环境中通过script标签引入）
if (typeof window !== 'undefined' && window.Vue) {
  install(window.Vue)
}

export default {
  install,
  ...components
}

// 导出各个组件供按需引入
export {
  DocumentViewer,
  FileUploader,
  UrlInput,
  ViewerControls
}
