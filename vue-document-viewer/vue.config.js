const path = require('path')

module.exports = {
  // 基础配置 - 添加需要转译的依赖
  transpileDependencies: [
    'mammoth',
    'jszip',
    'xlsx'
  ],
  
  // 公共路径
  publicPath: process.env.NODE_ENV === 'production' ? './' : '/',
  
  // 输出目录
  outputDir: 'dist',
  
  // 静态资源目录
  assetsDir: 'static',
  
  // 是否在构建生产包时生成 sourceMap 文件
  productionSourceMap: false,
  
  // 开发服务器配置
  devServer: {
    host: 'localhost',
    port: 8080,
    open: true,
    hot: true,
    overlay: {
      warnings: false,
      errors: true
    }
  },
  
  // Webpack配置
  configureWebpack: {
    // 路径别名
    resolve: {
      alias: {
        '@': path.resolve(__dirname, 'src'),
        'components': path.resolve(__dirname, 'src/components'),
        'utils': path.resolve(__dirname, 'src/utils'),
        'assets': path.resolve(__dirname, 'src/assets')
      }
    },
    
    // 性能优化
    optimization: {
      splitChunks: {
        chunks: 'all',
        cacheGroups: {
          // PDF.js 相关
          pdf: {
            name: 'pdf',
            test: /[\\/]node_modules[\\/]pdfjs-dist[\\/]/,
            priority: 10,
            reuseExistingChunk: true
          },
          // Office文档处理库
          office: {
            name: 'office',
            test: /[\\/]node_modules[\\/](mammoth|jszip|xlsx)[\\/]/,
            priority: 9,
            reuseExistingChunk: true
          },
          // Vue相关
          vue: {
            name: 'vue',
            test: /[\\/]node_modules[\\/](vue|vue-router)[\\/]/,
            priority: 8,
            reuseExistingChunk: true
          },
          // 第三方库
          vendors: {
            name: 'vendors',
            test: /[\\/]node_modules[\\/]/,
            priority: 5,
            reuseExistingChunk: true
          }
        }
      }
    }
  },
  
  // 链式配置
  chainWebpack: config => {
    // 修复HMR
    config.resolve.symlinks(true)
    
    
    // PDF.js worker 处理
    config.module
      .rule('worker')
      .test(/\.worker\.js$/)
      .use('worker-loader')
      .loader('worker-loader')
      .options({
        inline: 'fallback',
        filename: 'workers/[name].[contenthash:8].js'
      })
      .end()

    // 处理 PDF.js 的特殊需求
    config.module
      .rule('pdf-worker')
      .test(/pdf\.worker\.js$/)
      .use('file-loader')
      .loader('file-loader')
      .options({
        name: 'workers/pdf.worker.js'
      })
      .end()
    
    // 生产环境优化
    if (process.env.NODE_ENV === 'production') {
      // 移除 console
      config.optimization.minimizer('terser').tap(args => {
        if (args[0] && args[0].terserOptions) {
          args[0].terserOptions.compress = args[0].terserOptions.compress || {}
          args[0].terserOptions.compress.drop_console = true
          args[0].terserOptions.compress.drop_debugger = true
        }
        return args
      })
    }
  },
  
  // CSS相关配置
  css: {
    // 是否使用css分离插件 ExtractTextPlugin
    extract: process.env.NODE_ENV === 'production',
    
    // 开启 CSS source maps
    sourceMap: process.env.NODE_ENV === 'development'
  }
}