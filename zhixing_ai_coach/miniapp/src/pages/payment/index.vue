<template>
  <view class="page">
    <view class="order-card">
      <text class="order-title">方案定制</text>
      <text class="order-desc">AI + 8年教练经验 · 专属训练方案</text>
      <view class="price-row">
        <text class="price-symbol">¥</text>
        <text class="price-value">9.9</text>
      </view>
    </view>

    <view class="includes">
      <text class="section-title">方案包含</text>
      <view class="include-item">
        <image class="include-icon" src="/static/icons/document.svg" mode="widthFix" />
        <text>51 题精准问卷分析</text>
      </view>
      <view class="include-item">
        <image class="include-icon" src="/static/icons/ai.svg" mode="widthFix" />
        <text>AI 10 维决策树匹配</text>
      </view>
      <view class="include-item">
        <image class="include-icon" src="/static/icons/chart.svg" mode="widthFix" />
        <text>4 周完整训练计划</text>
      </view>
      <view class="include-item">
        <image class="include-icon" src="/static/icons/food.svg" mode="widthFix" />
        <text>个性化营养建议</text>
      </view>
      <view class="include-item">
        <image class="include-icon" src="/static/icons/shield.svg" mode="widthFix" />
        <text>关节保护专项方案</text>
      </view>
      <view class="include-item">
        <image class="include-icon" src="/static/icons/phone.svg" mode="widthFix" />
        <text>教练 24h 跟进支持</text>
      </view>
    </view>

    <button class="btn-pay" @tap="handlePay" :loading="paying">¥9.9 立即购买</button>
  </view>
</template>

<script>
export default {
  data() {
    return { paying: false, memberId: '' }
  },
  mounted() {
    const user = uni.getStorageSync('user')
    if (user) {
      const u = JSON.parse(user)
      this.createMemberAndOrder(u.nickname || '用户')
    }
  },
  methods: {
    async createMemberAndOrder(name) {
      uni.showLoading({ title: '初始化...' })
      try {
        const { default: api } = await import('@/api')
        const member = await api.createMember(name)
        this.memberId = member._id || member.id
        uni.hideLoading()
      } catch {
        uni.hideLoading()
        uni.showToast({ title: '创建失败', icon: 'none' })
      }
    },
    async handlePay() {
      if (!this.memberId) { uni.showToast({ title: '请重试', icon: 'none' }); return }
      this.paying = true
      try {
        const { default: api } = await import('@/api')
        const order = await api.createOrder(this.memberId)

        if (order.simulated) {
          uni.showToast({ title: '支付成功（模拟）' })
          setTimeout(() => uni.redirectTo({ url: `/pages/questionnaire/index?memberId=${this.memberId}` }), 1500)
        } else {
          uni.requestPayment({
            timeStamp: order.pay_params.timeStamp,
            nonceStr: order.pay_params.nonceStr,
            package: order.pay_params.package,
            signType: order.pay_params.signType,
            paySign: order.pay_params.paySign,
            success: () => {
              uni.showToast({ title: '支付成功' })
              setTimeout(() => uni.redirectTo({ url: `/pages/questionnaire/index?memberId=${this.memberId}` }), 1500)
            },
            fail: (err) => {
              if (err.errMsg?.includes('cancel')) uni.showToast({ title: '已取消支付', icon: 'none' })
              else uni.showToast({ title: '支付失败', icon: 'none' })
            },
          })
        }
      } catch (e) {
        uni.showToast({ title: e.message || '支付失败', icon: 'none' })
      } finally {
        this.paying = false
      }
    },
  },
}
</script>

<style scoped>
.page { padding: 40rpx 30rpx; }
.order-card {
  background: linear-gradient(135deg, #1a1a2e, #0f3460);
  border-radius: 20rpx;
  padding: 60rpx 40rpx;
  text-align: center;
  color: #fff;
}
.order-title { font-size: 40rpx; font-weight: 700; display: block; }
.order-desc { font-size: 26rpx; opacity: 0.8; margin-top: 12rpx; display: block; }
.price-row { margin-top: 30rpx; }
.price-symbol { font-size: 32rpx; }
.price-value { font-size: 80rpx; font-weight: 700; }
.includes { margin-top: 40rpx; }
.section-title { font-size: 32rpx; font-weight: 700; display: block; margin-bottom: 20rpx; }
.include-item {
  font-size: 28rpx;
  padding: 16rpx 0;
  border-bottom: 2rpx solid #f0f0f0;
  display: flex;
  align-items: center;
}
.include-icon {
  width: 36rpx;
  height: 36rpx;
  margin-right: 16rpx;
}
.btn-pay {
  margin-top: 60rpx;
  background: #e94560;
  color: #fff;
  border: none;
  border-radius: 50rpx;
  padding: 28rpx;
  font-size: 34rpx;
  font-weight: 600;
}
</style>

