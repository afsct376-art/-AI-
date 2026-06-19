const { cloud, db } = require('../database')
const config = require('../config')

exports.handle = async (event) => {
  const { code, nickname, avatarUrl } = event

  // 通过微信 jscode2session 换取 openid
  let openid = code
  if (config.WECHAT_APP_ID) {
    const res = await axios.get('https://api.weixin.qq.com/sns/jscode2session', {
      params: { appid: config.WECHAT_APP_ID, secret: config.WECHAT_APP_SECRET, js_code: code, grant_type: 'authorization_code' },
    })
    if (!res.data.openid) throw new Error(`微信登录失败: ${res.data.errmsg || '未知错误'}`)
    openid = res.data.openid
  }

  const { data: users } = await db.collection('users').where({ openid }).get()
  let user = users[0]

  if (!user) {
    const result = await db.collection('users').add({ data: { openid, nickname: nickname || '', avatarUrl: avatarUrl || '', isCoach: false, createdAt: db.serverDate(), updatedAt: db.serverDate() } })
    user = { _id: result.id, openid, nickname: nickname || '', avatarUrl: avatarUrl || '' }
  } else if (nickname) {
    await db.collection('users').doc(user._id).update({ data: { nickname, avatarUrl: avatarUrl || '', updatedAt: db.serverDate() } })
  }

  // 生成简单 token（用 _id 作为凭证）
  const token = Buffer.from(`${user._id}:${Date.now()}`).toString('base64')

  return { access_token: token, user: { id: user._id, nickname: user.nickname, avatarUrl: user.avatarUrl, isCoach: user.isCoach || false, createdAt: user.createdAt } }
}
