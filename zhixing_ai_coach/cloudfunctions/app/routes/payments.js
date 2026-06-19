const { db } = require('../database')
const { generateOrderNo, getPayParams } = require('../services/wechatPay')
const config = require('../config')

exports.handle = async (event, userId) => {
  const method = event.httpMethod

  if (method === 'POST' && event.path.endsWith('/create-order')) {
    const memberId = event.queryStringParameters?.member_id
    if (!memberId) throw new Error('缺少member_id')

    const { data: members } = await db.collection('members').where({ _id: memberId, userId }).get()
    if (!members.length) throw new Error('会员不存在')

    const orderNo = generateOrderNo()
    const result = await db.collection('orders').add({
      data: {
        userId, memberId, orderNo, totalFee: config.PLAN_PRICE,
        status: 'pending', prepayId: `mock_${orderNo}`, transactionId: '',
        createdAt: db.serverDate(), updatedAt: db.serverDate(),
      },
    })

    await db.collection('orders').doc(result.id).update({
      data: { status: 'paid', paidAt: db.serverDate() },
    })

    const payParams = getPayParams(`mock_${orderNo}`)
    return { order_id: result.id, order_no: orderNo, total_fee: config.PLAN_PRICE, prepay_id: `mock_${orderNo}`, pay_params: payParams, simulated: true }
  }

  if (method === 'GET' && event.path.endsWith('/orders')) {
    const { data } = await db.collection('orders').where({ userId }).orderBy('createdAt', 'desc').get()
    return data.map(o => ({ id: o._id, order_no: o.orderNo, total_fee: o.totalFee, status: o.status, paid_at: o.paidAt || null, created_at: o.createdAt }))
  }

  throw new Error('Not found')
}
