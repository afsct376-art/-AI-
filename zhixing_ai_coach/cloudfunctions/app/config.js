module.exports = {
  APP_NAME: '知行AI教练',
  APP_VERSION: '1.0.0',

  // 微信小程序配置（在CloudBase环境变量中设置）
  get WECHAT_APP_ID() {
    return process.env.WECHAT_APP_ID || ''
  },
  get WECHAT_APP_SECRET() {
    return process.env.WECHAT_APP_SECRET || ''
  },
  get WECHAT_MCH_ID() {
    return process.env.WECHAT_MCH_ID || ''
  },
  get WECHAT_API_KEY() {
    return process.env.WECHAT_API_KEY || ''
  },

  // AI API 配置
  get AI_API_KEY() {
    return process.env.AI_API_KEY || ''
  },
  AI_API_URL: 'https://api.deepseek.com/v1/chat/completions',
  AI_MODEL: 'deepseek-chat',

  // 方案定价（分）
  PLAN_PRICE: 990,

  // JWT 密钥
  get JWT_SECRET() {
    return process.env.JWT_SECRET || 'change-me'
  },
  JWT_EXPIRES: '7d',
}
