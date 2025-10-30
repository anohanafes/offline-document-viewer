module.exports = {
  presets: [
    [
      '@vue/cli-plugin-babel/preset',
      {
        useBuiltIns: 'entry',
        corejs: 3,
        targets: {
          browsers: ['> 1%', 'last 2 versions', 'not ie <= 8']
        }
      }
    ]
  ],
  plugins: [
    // 支持可选链操作符
    '@babel/plugin-transform-optional-chaining',
    // 支持空值合并操作符  
    '@babel/plugin-transform-nullish-coalescing-operator',
    // 支持类属性
    '@babel/plugin-transform-class-properties',
    // 支持私有方法和属性
    '@babel/plugin-transform-private-methods',
    // 支持私有属性
    '@babel/plugin-transform-private-property-in-object',
    // 支持静态类属性
    '@babel/plugin-transform-class-static-block'
  ],
  env: {
    development: {
      plugins: [
        // 开发环境插件
      ]
    },
    production: {
      plugins: [
        // 生产环境插件
      ]
    }
  }
}