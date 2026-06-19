const { ENV_ID } = require('./config')

try {
  if (wx.cloud) {
    wx.cloud.init({ env: ENV_ID, traceUser: true })
  }
} catch (e) {}

// 统一调用云函数
export function callFunction(path, method = 'GET', body = {}) {
  return new Promise((resolve, reject) => {
    if (!wx.cloud) {
      reject(new Error('请在微信小程序中运行'))
      return
    }

    // 构建符合云函数 event 格式的数据
    const data = {
      path,
      httpMethod: method,
      headers: {},
      queryStringParameters: {},
      body: method === 'GET' ? '' : JSON.stringify(body),
    }

    // 对登录等特殊接口，直接传参（auth.handle 直接从 event 读字段）
    if (path === '/api/auth/login') {
      data.code = body.code
      data.nickname = body.nickname || ''
      data.avatarUrl = body.avatarUrl || ''
    }

    wx.cloud.callFunction({
      name: 'app',
      data,
      success: (res) => {
        const result = res.result
        if (result.code === 0) {
          resolve(result.data)
        } else {
          reject(new Error(result.message || '请求失败'))
        }
      },
      fail: (err) => {
        reject(new Error(err.errMsg || '网络异常'))
      },
    })
  })
}

// 带认证的调用（自动带 token）
export function authCall(path, method = 'GET', body = {}) {
  const token = uni.getStorageSync('token') || ''
  return new Promise((resolve, reject) => {
    if (!wx.cloud) {
      reject(new Error('请在微信小程序中运行'))
      return
    }

    const data = {
      path,
      httpMethod: method,
      headers: { Authorization: `Bearer ${token}` },
      queryStringParameters: {},
      body: method === 'GET' ? '' : JSON.stringify(body),
    }

    // 对 createMember，直接传 name
    if (path === '/api/users/members' && method === 'POST') {
      data.name = body.name
    }

    // 对 create-order，把 member_id 放入 queryStringParameters
    if (path === '/api/payments/create-order') {
      data.queryStringParameters = { member_id: body.member_id }
    }

    wx.cloud.callFunction({
      name: 'app',
      data,
      success: (res) => {
        const result = res.result
        if (result.code === 0) {
          resolve(result.data)
        } else if (result.message === '未登录') {
          uni.removeStorageSync('token')
          uni.navigateTo({ url: '/pages/index/index' })
          reject(new Error('登录已过期'))
        } else {
          reject(new Error(result.message || '请求失败'))
        }
      },
      fail: (err) => {
        reject(new Error(err.errMsg || '网络异常'))
      },
    })
  })
}

// 便捷 API
export default {
  health: () => callFunction('/api/health'),

  login: (code, nickname = '', avatarUrl = '') =>
    callFunction('/api/auth/login', 'POST', { code, nickname, avatarUrl }),

  createMember: (name) =>
    authCall('/api/users/members', 'POST', { name }),

  listMembers: () =>
    authCall('/api/users/members'),

  submitQuestionnaire: (memberId, rawData) =>
    authCall('/api/questionnaire/submit', 'POST', { member_id: memberId, raw_data: rawData }),

  getQuestionnaire: (memberId) =>
    authCall(`/api/questionnaire/${memberId}`),

  getPlan: (memberId) =>
    authCall(`/api/plans/${memberId}`),

  generatePlan: (memberId, coachNotes = '') =>
    authCall('/api/plans/generate', 'POST', { member_id: memberId, coach_notes: coachNotes }),

  createOrder: (memberId) =>
    authCall('/api/payments/create-order', 'POST', { member_id: memberId }),

  listOrders: () =>
    authCall('/api/payments/orders'),
}
