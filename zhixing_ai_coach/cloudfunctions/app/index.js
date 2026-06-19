const auth = require('./routes/auth')
const members = require('./routes/members')
const questionnaire = require('./routes/questionnaire')
const plans = require('./routes/plans')
const payments = require('./routes/payments')

// 从 token 解析 userId
function parseToken(token) {
  if (!token) return ''
  try {
    const buf = Buffer.from(token, 'base64').toString()
    return buf.split(':')[0]
  } catch { return '' }
}

// CloudBase 云函数入口
exports.main = async (event, context) => {
  console.log('[REQUEST]', event.httpMethod, event.path)

  try {
    const path = event.path || ''
    const method = event.httpMethod || 'GET'

    // 解析 token
    const authHeader = event.headers?.['authorization'] || event.headers?.['Authorization'] || ''
    const token = authHeader.replace('Bearer ', '')
    const userId = parseToken(token)

    let result

    // 路由分发
    if (path === '/api/auth/login' && method === 'POST') {
      result = await auth.handle(event)
    } else if (path.startsWith('/api/users/members')) {
      if (!userId) throw new Error('未登录')
      result = await members.handle(event, userId)
    } else if (path.startsWith('/api/questionnaire/')) {
      if (!userId) throw new Error('未登录')
      result = await questionnaire.handle(event, userId)
    } else if (path.startsWith('/api/plans/')) {
      if (!userId) throw new Error('未登录')
      result = await plans.handle(event, userId)
    } else if (path.startsWith('/api/payments/')) {
      if (!userId) throw new Error('未登录')
      result = await payments.handle(event, userId)
    } else if (path === '/api/health') {
      result = { status: 'ok', version: '1.0.0' }
    } else {
      throw new Error('Not found')
    }

    return { code: 0, data: result, message: 'ok' }
  } catch (e) {
    console.error('[ERROR]', e.message)
    return { code: -1, message: e.message }
  }
}
