<template>
  <view class="home">
    <view class="hero">
      <text class="hero-title">知行AI教练</text>
      <text class="hero-subtitle">AI + 真人教练 · 你的专属训练方案</text>
      <text class="hero-desc">已完成 8 份定制方案 · 限时体验价 ¥9.9</text>
      <button class="btn-start" @tap="startNow">立即定制</button>
    </view>

    <!-- 快捷入口 -->
    <view class="quick-actions">
      <view class="action-btn" @tap="navigateTo('/pages/plan/index')">
        <image class="action-icon" src="/static/icons/target.svg" />
        <text>我的方案</text>
      </view>
      <view class="action-btn" @tap="navigateTo('/pages/questionnaire/index')">
        <image class="action-icon" src="/static/icons/ai.svg" />
        <text>快速诊断</text>
      </view>
    </view>

    <view class="features">
      <view class="section-title">为什么选择知行AI教练</view>
      <view class="feature-grid">
        <view class="feature-card">
          <image class="feature-icon" src="/static/icons/ai.svg" mode="widthFix" />
          <text class="feature-title">AI 智能分析</text>
          <text class="feature-desc">10 维决策树 + 51 题精准问卷</text>
        </view>
        <view class="feature-card">
          <image class="feature-icon" src="/static/icons/coach.svg" mode="widthFix" />
          <text class="feature-title">教练微调</text>
          <text class="feature-desc">8 年经验教练把关，拒绝通用模板</text>
        </view>
        <view class="feature-card">
          <image class="feature-icon" src="/static/icons/target.svg" mode="widthFix" />
          <text class="feature-title">7 种会员类型</text>
          <text class="feature-desc">零基础到同行教练，精准匹配</text>
        </view>
        <view class="feature-card">
          <image class="feature-icon" src="/static/icons/document.svg" mode="widthFix" />
          <text class="feature-title">完整交付</text>
          <text class="feature-desc">Excel 方案 + 动作详解 + 营养建议</text>
        </view>
      </view>
    </view>

    <view class="member-types">
      <view class="section-title">7 种会员类型</view>
      <view class="type-list">
        <view class="type-item" v-for="t in memberTypes" :key="t.id">
          <text class="type-tag">{{ t.id }}</text>
          <view class="type-info">
            <text class="type-name">{{ t.name }}</text>
            <text class="type-desc">{{ t.desc }}</text>
          </view>
        </view>
      </view>
    </view>

    <view class="cases">
      <view class="section-title">真实案例</view>
      <view class="case-card" v-for="c in cases" :key="c.name">
        <text class="case-type">{{ c.type }}</text>
        <text class="case-name">{{ c.name }}</text>
        <text class="case-desc">{{ c.desc }}</text>
      </view>
    </view>

    <!-- 悬浮客服按钮 -->
    <view class="floating-chat" @tap="contactService">
      <image class="chat-icon" src="/static/icons/phone.svg" />
    </view>
  </view>
</template>

<script>
export default {
  data() {
    return {
      memberTypes: [
        { id: 'A', name: '渐进入门型', desc: '零基础，不知道怎么开始' },
        { id: 'B', name: '重建信心型', desc: '停训后信心消退，需要重新开始' },
        { id: 'C', name: '进阶突破型', desc: '中级遇到瓶颈，需要突破' },
        { id: 'D', name: '精细优化型', desc: '高级+专项需求，精益求精' },
        { id: 'E', name: '同行教练型', desc: '健身教练，专业交流' },
        { id: 'F', name: '老会员进阶型', desc: '已有信任，直接进阶' },
        { id: 'G', name: '高压工作型', desc: '工作压力大，做减法更重要' },
      ],
      cases: [
        { name: 'Joe', type: 'TYPE_C', desc: 'DUP+RPE+器械突破+季度全休' },
        { name: '蜡笔小杨', type: 'TYPE_B', desc: '继续旅程+双路径夜班方案' },
        { name: '小冯', type: 'TYPE_G', desc: 'ABC三方协作+动作库前置+成绩单封装' },
      ],
    }
  },
  methods: {
    startNow() {
      const token = uni.getStorageSync('token')
      if (token) {
        uni.navigateTo({ url: '/pages/questionnaire/index' })
      } else {
        this.wechatLogin()
      }
    },
    contactService() {
      uni.showToast({ title: '客服功能建设中', icon: 'none' })
    },
    navigateTo(url) {
      uni.navigateTo({ url })
    },
    async wechatLogin() {
      uni.showLoading({ title: '登录中...' })
      try {
        const { code } = await uni.login({ provider: 'weixin' })
        const { default: api } = await import('@/api')
        const userData = await api.login(code)
        uni.setStorageSync('token', userData.access_token)
        uni.setStorageSync('user', JSON.stringify(userData.user))
        uni.navigateTo({ url: '/pages/payment/index' })
      } catch (e) {
        uni.showToast({ title: '请允许微信登录', icon: 'none' })
      }
      uni.hideLoading()
    },
  },
}
</script>

<style scoped>
.home { padding-bottom: 40rpx; }
.hero {
  background: linear-gradient(135deg, var(--dark-color) 0%, #16213e 100%);
  padding: 80rpx 40rpx 100rpx;
  text-align: center;
  color: #fff;
  border-bottom-left-radius: 40rpx;
  border-bottom-right-radius: 40rpx;
}
.hero-title { font-size: 52rpx; font-weight: 700; display: block; }
.hero-subtitle { font-size: 30rpx; opacity: 0.85; margin-top: 16rpx; display: block; }
.hero-desc { font-size: 26rpx; opacity: 0.7; margin-top: 12rpx; display: block; }
.btn-start {
  margin-top: 40rpx;
  background: var(--primary-color);
  color: #fff;
  border: none;
  border-radius: var(--radius-lg);
  padding: 24rpx 80rpx;
  font-size: 32rpx;
  font-weight: 600;
  box-shadow: 0 10rpx 20rpx rgba(233, 69, 96, 0.3);
}
.section-title {
  font-size: 34rpx;
  font-weight: 700;
  padding: 40rpx 30rpx 20rpx;
  color: #1a1a2e;
}
.feature-grid { display: flex; flex-wrap: wrap; padding: 0 20rpx; }
.feature-card {
  width: 45%;
  background: #fff;
  border-radius: var(--radius-lg);
  margin: 10rpx;
  padding: 30rpx 24rpx;
  box-shadow: var(--shadow-base);
}
.feature-icon { width: 48rpx; height: 48rpx; display: block; }
.feature-title { font-size: 30rpx; font-weight: 600; margin-top: 12rpx; display: block; }
.feature-desc { font-size: 24rpx; color: #666; margin-top: 8rpx; display: block; }
.type-list { padding: 0 30rpx; }
.type-item {
  display: flex;
  align-items: center;
  background: #fff;
  border-radius: 12rpx;
  padding: 24rpx;
  margin-bottom: 16rpx;
  box-shadow: 0 2rpx 8rpx rgba(0,0,0,0.04);
}
.type-tag {
  background: #e94560;
  color: #fff;
  font-size: 24rpx;
  font-weight: 700;
  padding: 8rpx 16rpx;
  border-radius: 8rpx;
  margin-right: 20rpx;
}
.type-info { flex: 1; }
.type-name { font-size: 30rpx; font-weight: 600; display: block; }
.type-desc { font-size: 24rpx; color: #666; margin-top: 4rpx; display: block; }
.quick-actions {
  display: flex;
  justify-content: space-around;
  padding: 30rpx;
  background: #fff;
  margin: -40rpx 30rpx 20rpx;
  border-radius: 20rpx;
  box-shadow: 0 4rpx 16rpx rgba(0,0,0,0.08);
}
.action-btn { display: flex; flex-direction: column; align-items: center; }
.action-icon { width: 64rpx; height: 64rpx; margin-bottom: 8rpx; }
.floating-chat {
  position: fixed;
  bottom: 100rpx;
  right: 40rpx;
  width: 100rpx;
  height: 100rpx;
  background: #e94560;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  box-shadow: 0 4rpx 20rpx rgba(233, 69, 96, 0.4);
}
.chat-icon { width: 50rpx; height: 50rpx; }
.case-card {
  background: #fff;
  border-left: 8rpx solid #e94560;
  padding: 24rpx;
  margin-bottom: 12rpx;
  border-radius: 8rpx;
}
.case-type { font-size: 22rpx; color: #e94560; font-weight: 600; display: block; }
.case-name { font-size: 30rpx; font-weight: 600; display: block; margin-top: 4rpx; }
.case-desc { font-size: 24rpx; color: #666; display: block; margin-top: 4rpx; }
</style>

