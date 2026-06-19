const crypto = require('crypto')

function generateOrderNo() {
  const ts = Date.now()
  const rand = Math.random().toString().slice(2, 8)
  return `ZX${ts}${rand}`
}

function getPayParams(prepayId) {
  const nonceStr = crypto.randomBytes(16).toString('hex')
  const timeStamp = Math.floor(Date.now() / 1000).toString()
  const packageStr = `prepay_id=${prepayId}`

  return {
    timeStamp,
    nonceStr,
    package: packageStr,
    signType: 'MD5',
    paySign: 'simulated-sign',
  }
}

module.exports = { generateOrderNo, getPayParams }
