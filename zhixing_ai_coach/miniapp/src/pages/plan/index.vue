<template>
  <view class="page">
    <view v-if="loading" class="loading-state">
      <text class="loading-text">方案生成中...</text>
      <text class="loading-desc">AI 正在根据你的信息定制方案，请稍候...</text>
    </view>

    <view v-else-if="error" class="error-state">
      <text class="error-text">生成失败</text>
      <text class="error-desc">{{ error }}</text>
      <button class="btn-retry" @tap="loadPlan">重新生成</button>
    </view>

    <template v-else>
      <view class="plan-header" :style="{ background: headerColor }">
        <text class="plan-type">{{ planTypeName }}</text>
        <text class="plan-subtitle">{{ planSubtitle }}</text>
        <view class="plan-stats">
          <view class="stat">
            <text class="stat-value">{{ plan.weekly_frequency }}</text>
            <text class="stat-label">次/周</text>
          </view>
          <view class="stat">
            <text class="stat-value">{{ plan.session_duration }}</text>
            <text class="stat-label">分钟/次</text>
          </view>
          <view class="stat">
            <text class="stat-value">{{ plan.split_type }}</text>
            <text class="stat-label">分化方式</text>
          </view>
        </view>
      </view>

      <view class="plan-content">
        <rich-text :nodes="renderedContent"></rich-text>
      </view>

      <button class="btn-contact" open-type="contact">联系教练</button>
    </template>
  </view>
</template>

<script>
const TYPE_CONFIG = {
  TYPE_A: { name: '渐进入门型', subtitle: '第一周——只做这三件事', color: 'linear-gradient(135deg, #667eea, #764ba2)' },
  TYPE_B: { name: '重建信心型', subtitle: '你不是重新开始——你是继续旅程', color: 'linear-gradient(135deg, #f093fb, #f5576c)' },
  TYPE_C: { name: '进阶突破型', subtitle: '今天练什么——找到今天的主题', color: 'linear-gradient(135deg, #4facfe, #00f2fe)' },
  TYPE_D: { name: '精细优化型', subtitle: '力量服务耐力', color: 'linear-gradient(135deg, #43e97b, #38f9d7)' },
  TYPE_E: { name: '同行教练型', subtitle: '你是教练——这份方案写给你的', color: 'linear-gradient(135deg, #fa709a, #fee140)' },
  TYPE_F: { name: '老会员进阶型', subtitle: '周期化进阶', color: 'linear-gradient(135deg, #a18cd1, #fbc2eb)' },
  TYPE_G: { name: '高压工作型', subtitle: '你的第一个月——数据不会骗人', color: 'linear-gradient(135deg, #1a1a2e, #16213e)' },
}

export default {
  data() {
    return {
      memberId: '',
      plan: null,
      loading: true,
      error: '',
    }
  },
  onLoad(query) {
    this.memberId = query.memberId || ''
    this.loadPlan()
  },
  mounted() {
    if (!this.memberId) this.loadPlan()
  },
  computed: {
    headerColor() {
      return TYPE_CONFIG[this.plan?.plan_type]?.color || 'linear-gradient(135deg, #1a1a2e, #0f3460)'
    },
    planTypeName() {
      return TYPE_CONFIG[this.plan?.plan_type]?.name || this.plan?.plan_type || ''
    },
    planSubtitle() {
      return TYPE_CONFIG[this.plan?.plan_type]?.subtitle || ''
    },
    renderedContent() {
      if (!this.plan?.raw_content) return ''
      let html = this.plan.raw_content
        .replace(/### /g, '<h3>')
        .replace(/## /g, '<h2>')
        .replace(/# /g, '<h1>')
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/\n/g, '<br/>')
      return html
    },
  },
  methods: {
    async loadPlan() {
      if (!this.memberId) { this.error = '缺少会员信息'; this.loading = false; return }
      this.loading = true
      this.error = ''
      try {
        const { default: api } = await import('@/api')
        const plan = await api.getPlan(this.memberId)
        this.plan = plan
        if (plan.status === 'generating') {
          setTimeout(() => this.loadPlan(), 3000)
        }
      } catch (e) {
        this.error = e.message
      } finally {
        this.loading = false
      }
    },
  },
}
</script>

<style scoped>
.page { min-height: 100vh; }
.loading-state, .error-state {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 200rpx 60rpx;
  text-align: center;
}
.loading-text { font-size: 36rpx; font-weight: 700; color: #1a1a2e; }
.loading-desc { font-size: 28rpx; color: #999; margin-top: 16rpx; }
.error-text { font-size: 36rpx; font-weight: 700; color: #e94560; }
.error-desc { font-size: 28rpx; color: #999; margin-top: 16rpx; }
.btn-retry { margin-top: 30rpx; background: #e94560; color: #fff; border: none; border-radius: 12rpx; padding: 20rpx 60rpx; }
.plan-header {
  padding: 60rpx 40rpx;
  text-align: center;
  color: #fff;
}
.plan-type { font-size: 40rpx; font-weight: 700; display: block; }
.plan-subtitle { font-size: 28rpx; opacity: 0.85; margin-top: 12rpx; display: block; }
.plan-stats { display: flex; justify-content: center; margin-top: 30rpx; }
.stat { margin: 0 30rpx; text-align: center; }
.stat-value { font-size: 36rpx; font-weight: 700; display: block; }
.stat-label { font-size: 24rpx; opacity: 0.7; display: block; margin-top: 4rpx; }
.plan-content { padding: 30rpx; }
.btn-contact {
  margin: 40rpx 30rpx;
  background: #1a1a2e;
  color: #fff;
  border: none;
  border-radius: 50rpx;
  padding: 24rpx;
  font-size: 30rpx;
}
</style>

