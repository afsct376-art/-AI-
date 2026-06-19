<template>
  <view class="page">
    <view class="user-info">
      <view class="avatar">{{ (user?.nickname || 'U').charAt(0) }}</view>
      <text class="nickname">{{ user?.nickname || '用户' }}</text>
    </view>

    <view class="section">
      <text class="section-title">我的方案</text>
      <view class="member-card" v-for="m in members" :key="m.id" @tap="goToPlan(m)">
        <view class="member-tag" :class="m.member_type?.toLowerCase()">{{ m.member_type || '待分类' }}</view>
        <view class="member-body">
          <text class="member-name">{{ m.name }}</text>
          <text class="member-status">{{ statusText(m) }}</text>
        </view>
        <text class="member-arrow"></text>
      </view>
      <view v-if="members.length === 0" class="empty">
        <text>还没有方案</text>
        <button class="btn-start" @tap="goHome">立即定制</button>
      </view>
    </view>

    <view class="section">
      <text class="section-title">购买记录</text>
      <view class="order-card" v-for="o in orders" :key="o.id">
        <view class="order-top">
          <text class="order-no">{{ o.order_no }}</text>
          <text class="order-status" :class="o.status">{{ o.status === 'paid' ? '已支付' : '待支付' }}</text>
        </view>
        <text class="order-price">¥{{ (o.total_fee / 100).toFixed(1) }}</text>
        <text class="order-time">{{ o.created_at }}</text>
      </view>
    </view>

    <button class="btn-logout" @tap="handleLogout">退出登录</button>
  </view>
</template>

<script>
export default {
  data() {
    return { user: null, members: [], orders: [] }
  },
  mounted() {
    const u = uni.getStorageSync('user')
    if (u) this.user = JSON.parse(u)
    this.loadData()
  },
  methods: {
    async loadData() {
      uni.showLoading({ title: '加载中...' })
      try {
        const { default: api } = await import('@/api')
        const [members, orders] = await Promise.all([
          api.listMembers(),
          api.listOrders(),
        ])
        this.members = members.members || []
        this.orders = orders
      } catch (e) {
        uni.showToast({ title: '加载失败', icon: 'none' })
      }
      uni.hideLoading()
    },
    statusText(m) {
      const map = { pending: '待填写问卷', questionnaire_done: '问卷已提交', generating: '方案生成中', completed: '方案已就绪' }
      return map[m.status] || m.status
    },
    goToPlan(m) {
      if (m.status === 'pending') uni.navigateTo({ url: `/pages/questionnaire/index?memberId=${m.id}` })
      else if (m.status === 'completed') uni.navigateTo({ url: `/pages/plan/index?memberId=${m.id}` })
    },
    goHome() { uni.switchTab({ url: '/pages/index/index' }) },
    handleLogout() {
      uni.showModal({
        title: '确认退出',
        content: '退出后需要重新登录',
        success: (res) => { if (res.confirm) { uni.removeStorageSync('token'); uni.removeStorageSync('user'); uni.reLaunch({ url: '/pages/index/index' }) } },
      })
    },
  },
}
</script>

<style scoped>
.page { padding: 30rpx; }
.user-info {
  display: flex;
  align-items: center;
  padding: 40rpx 0;
}
.avatar {
  width: 100rpx;
  height: 100rpx;
  border-radius: 50%;
  background: linear-gradient(135deg, #e94560, #0f3460);
  color: #fff;
  font-size: 40rpx;
  font-weight: 700;
  text-align: center;
  line-height: 100rpx;
  margin-right: 24rpx;
}
.nickname { font-size: 34rpx; font-weight: 600; }
.section { margin-top: 40rpx; }
.section-title { font-size: 32rpx; font-weight: 700; display: block; margin-bottom: 16rpx; }
.member-card {
  display: flex;
  align-items: center;
  background: #fff;
  border-radius: 16rpx;
  padding: 24rpx;
  margin-bottom: 12rpx;
  box-shadow: 0 2rpx 8rpx rgba(0,0,0,0.04);
}
.member-tag {
  font-size: 20rpx;
  font-weight: 600;
  padding: 6rpx 14rpx;
  border-radius: 8rpx;
  margin-right: 16rpx;
  background: #f0f0f0;
  color: #666;
}
.member-tag.type_a { background: #e8f4fd; color: #1976d2; }
.member-tag.type_g { background: #fce4ec; color: #c62828; }
.member-body { flex: 1; }
.member-name { font-size: 30rpx; font-weight: 600; display: block; }
.member-status { font-size: 24rpx; color: #999; display: block; margin-top: 4rpx; }
.member-arrow { font-size: 40rpx; color: #ccc; }
.empty { text-align: center; padding: 60rpx; color: #999; font-size: 28rpx; }
.btn-start { margin-top: 20rpx; background: #e94560; color: #fff; border: none; border-radius: 12rpx; padding: 16rpx 60rpx; font-size: 28rpx; display: inline-block; }
.order-card {
  background: #fff;
  border-radius: 12rpx;
  padding: 20rpx;
  margin-bottom: 12rpx;
}
.order-top { display: flex; justify-content: space-between; }
.order-no { font-size: 24rpx; color: #999; }
.order-status { font-size: 24rpx; font-weight: 600; }
.order-status.paid { color: #4caf50; }
.order-status.pending { color: #ff9800; }
.order-price { font-size: 36rpx; font-weight: 700; display: block; margin-top: 8rpx; }
.order-time { font-size: 22rpx; color: #ccc; display: block; margin-top: 4rpx; }
.btn-logout {
  margin-top: 60rpx;
  background: none;
  border: 2rpx solid #e0e0e0;
  border-radius: 12rpx;
  padding: 20rpx;
  font-size: 28rpx;
  color: #999;
}
</style>

