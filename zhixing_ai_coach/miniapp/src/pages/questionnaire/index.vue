<template>
  <view class="page">
    <view class="progress-bar">
      <view class="progress-fill" :style="{ width: progress + '%' }"></view>
    </view>
    <text class="progress-text">{{ currentStep }}/{{ totalSteps }} 已完成 {{ progress }}%</text>

    <view class="form-section" v-for="(section, si) in sections" :key="si" v-show="si === currentStepIndex">
      <text class="section-title">{{ section.title }}</text>
      <text class="section-desc">{{ section.desc }}</text>

      <view class="form-group" v-for="(q, qi) in section.questions" :key="qi">
        <text class="question-label">{{ q.label }}</text>
        <text class="question-hint" v-if="q.hint">{{ q.hint }}</text>

        <!-- Text input -->
        <van-field v-if="q.type === 'text'" v-model="answers[q.key]" :label="q.label" :placeholder="q.placeholder || '请输入'" />

        <!-- Textarea -->
        <van-field v-if="q.type === 'textarea'" v-model="answers[q.key]" type="textarea" :label="q.label" :placeholder="q.placeholder || '请详细描述'" autosize />

        <!-- Radio -->
        <view v-if="q.type === 'radio'" class="options">
          <text class="question-label">{{ q.label }}</text>
          <van-radio-group :value="answers[q.key]" @change="(e) => answers[q.key] = e.detail">
            <van-radio v-for="opt in q.options" :key="opt" :name="opt" custom-class="option-item">{{ opt }}</van-radio>
          </van-radio-group>
        </view>

        <!-- Checkbox -->
        <view v-if="q.type === 'checkbox'" class="options">
          <text class="question-label">{{ q.label }}</text>
          <van-checkbox-group :value="answers[q.key] || []" @change="(e) => answers[q.key] = e.detail">
            <van-checkbox v-for="opt in q.options" :key="opt" :name="opt" custom-class="option-item">{{ opt }}</van-checkbox>
          </van-checkbox-group>
        </view>

        <!-- Select -->
        <picker v-if="q.type === 'select'" :value="answers[q.key] || ''" :range="q.options" @change="(e) => answers[q.key] = q.options[e.detail.value]">
          <view class="picker">{{ answers[q.key] || '请选择' }}</view>
        </picker>
      </view>
    </view>

    <view class="form-nav">
      <button v-if="currentStepIndex > 0" class="btn-prev" @tap="prevStep">上一页</button>
      <button v-if="currentStepIndex < totalSteps - 1" class="btn-next" @tap="nextStep">下一步</button>
      <button v-if="currentStepIndex === totalSteps - 1" class="btn-submit" @tap="submitForm">提交问卷</button>
    </view>
  </view>
</template>

<script>
export default {
  data() {
    return {
      memberId: '',
      currentStepIndex: 0,
      answers: {},
      sections: [
        {
          title: '基本信息',
          desc: '先认识一下你',
          questions: [
            { key: 'Q1', label: '你的姓名', type: 'text' },
            { key: 'Q2', label: '你的年龄', type: 'select', options: ['18岁以下', '18-25岁', '26-30岁', '31-35岁', '36-40岁', '41-45岁', '46-50岁', '50岁以上'] },
            { key: 'Q3', label: '你的性别', type: 'radio', options: ['男', '女'] },
            { key: 'Q4', label: '你的身高（cm）', type: 'text', placeholder: '例：175' },
            { key: 'Q5', label: '你的体重（kg）', type: 'text', placeholder: '例：70' },
          ],
        },
        {
          title: '训练经验',
          desc: '了解你的训练背景',
          questions: [
            { key: 'Q6', label: '你的训练经验水平', type: 'radio', options: ['完全新手', '初学者（<6个月）', '中级（6个月-2年）', '高级（2年以上）', '曾经练过，现在重新开始'] },
            { key: 'Q7', label: '每周能训练几次？', type: 'radio', options: ['1次', '2次', '3次', '4次', '5次'] },
            { key: 'Q8', label: '每次训练可用多长时间？', type: 'radio', options: ['30分钟', '30-45分钟', '45-60分钟', '60-90分钟'] },
          ],
        },
        {
          title: '训练目标',
          desc: '你想要达成什么？',
          questions: [
            { key: 'Q9', label: '主要训练目标', type: 'checkbox', options: ['增肌', '减脂', '塑形', '力量提升', '综合体能', '保持健康', '学会自己做计划'] },
            { key: 'Q10', label: '请描述你的具体目标', type: 'textarea', placeholder: '写得越具体，方案越精准' },
          ],
        },
        {
          title: '器械条件',
          desc: '你在哪里练、用什么练',
          questions: [
            { key: 'Q11', label: '训练地点', type: 'radio', options: ['商业健身房', '铁馆', '酒店健身房', '公司健身房', '小区健身房', '家庭健身', '纯徒手户外'] },
            { key: 'Q12', label: '可用器械（可多选）', type: 'checkbox', options: ['哑铃', '杠铃', '龙门架', '史密斯机', '弹力带', '泡沫轴', 'TRX', '单杠'] },
          ],
        },
        {
          title: '运动损伤',
          desc: '安全第一，请如实填写',
          questions: [
            { key: 'Q13', label: '是否有运动伤病史？', type: 'radio', options: ['无', '有（请说明）'] },
            { key: 'Q14', label: '目前是否有身体疼痛或不适？', type: 'textarea', placeholder: '部位+症状+什么动作触发' },
          ],
        },
        {
          title: '心理与偏好',
          desc: '你的心态决定了方案的方向',
          questions: [
            { key: 'Q19f', label: '目前对健身的心态更接近哪种？', type: 'radio', options: ['很想开始但一直没动起来', '之前练得好，停了一段时间不知道怎么重新开始', '有基础但遇到了瓶颈', '工作太忙/压力大，练了又断有点内疚', '一直在规律训练想要继续进阶', '健身教练来找专业交流'] },
          ],
        },
        {
          title: '工作与生活',
          desc: '你的生活节奏影响训练节奏',
          questions: [
            { key: 'Q19c', label: '工作压力程度？', type: 'radio', options: ['轻松', '正常', '偏大', '极大（经常加班）'] },
            { key: 'Q19e', label: '单程通勤时间？', type: 'radio', options: ['30分钟以内', '30-60分钟', '60-90分钟', '90分钟以上'] },
            { key: 'Q25', label: '睡眠情况？', type: 'radio', options: ['好（7-8小时）', '一般（6-7小时）', '差（<6小时）'] },
          ],
        },
      ],
      totalSteps: 7,
    }
  },
  onLoad(query) {
    this.memberId = query.memberId || ''
  },
  computed: {
    progress() {
      return Math.round(((this.currentStepIndex + 1) / this.totalSteps) * 100)
    },
    currentStep() {
      return this.currentStepIndex + 1
    },
  },
  methods: {
    toggleCheckbox(key, opt) {
      if (!this.answers[key]) this.answers[key] = []
      const idx = this.answers[key].indexOf(opt)
      if (idx > -1) this.answers[key].splice(idx, 1)
      else this.answers[key].push(opt)
    },
    nextStep() {
      const qs = this.sections[this.currentStepIndex].questions
      for (const q of qs) {
        if (q.type !== 'textarea' && q.type !== 'checkbox' && !this.answers[q.key]) {
          uni.showToast({ title: '请先完成当前部分', icon: 'none' })
          return
        }
      }
      this.currentStepIndex++
      uni.pageScrollTo({ scrollTop: 0, duration: 200 })
    },
    prevStep() {
      this.currentStepIndex--
      uni.pageScrollTo({ scrollTop: 0, duration: 200 })
    },
    async submitForm() {
      if (!this.memberId) { uni.showToast({ title: '请先购买方案', icon: 'none' }); return }
      uni.showLoading({ title: '提交中...' })
      try {
        const { default: api } = await import('@/api')
        await api.submitQuestionnaire(this.memberId, this.answers)
        uni.hideLoading()
        uni.showToast({ title: '提交成功' })
        await api.generatePlan(this.memberId)
        setTimeout(() => {
          uni.redirectTo({ url: `/pages/plan/index?memberId=${this.memberId}` })
        }, 1000)
      } catch (e) {
        uni.hideLoading()
        uni.showToast({ title: e.message || '提交失败', icon: 'none' })
      }
    },
  },
}
</script>

<style scoped>
.page { padding: 30rpx; }
.progress-bar {
  height: 8rpx;
  background: #e0e0e0;
  border-radius: 4rpx;
  overflow: hidden;
}
.progress-fill {
  height: 100%;
  background: linear-gradient(90deg, #e94560, #0f3460);
  border-radius: 4rpx;
  transition: width 0.3s;
}
.progress-text { font-size: 24rpx; color: #999; margin-top: 8rpx; display: block; text-align: center; }
.form-section { margin-top: 30rpx; }
.section-title { font-size: 36rpx; font-weight: 700; color: #1a1a2e; display: block; }
.section-desc { font-size: 26rpx; color: #999; margin-top: 8rpx; display: block; }
.form-group { margin-top: 32rpx; }
.question-label { font-size: 30rpx; font-weight: 600; display: block; }
.question-hint { font-size: 24rpx; color: #999; margin-top: 6rpx; display: block; }
.input, .textarea, .picker {
  width: 100%;
  border: 2rpx solid #e0e0e0;
  border-radius: 12rpx;
  padding: 20rpx;
  margin-top: 12rpx;
  font-size: 28rpx;
  box-sizing: border-box;
  background: #fff;
}
.textarea { height: 160rpx; }
.options { margin-top: 12rpx; }
.option {
  display: flex;
  align-items: center;
  padding: 20rpx;
  background: #fff;
  border: 2rpx solid #e8e8e8;
  border-radius: 12rpx;
  margin-bottom: 12rpx;
  transition: all 0.2s;
}
.option.active { border-color: #e94560; background: #fff5f5; }
.option-dot {
  width: 32rpx;
  height: 32rpx;
  border: 3rpx solid #ccc;
  border-radius: 50%;
  margin-right: 16rpx;
}
.option-dot.filled { border-color: #e94560; background: #e94560; }
.option-check {
  width: 32rpx;
  height: 32rpx;
  border: 3rpx solid #ccc;
  border-radius: 6rpx;
  margin-right: 16rpx;
  text-align: center;
  line-height: 32rpx;
  font-size: 24rpx;
  color: transparent;
}
.option-check.checked { border-color: #e94560; background: #e94560; color: #fff; }
.option-text { font-size: 28rpx; }
.form-nav {
  display: flex;
  justify-content: space-between;
  margin-top: 40rpx;
  padding-bottom: 60rpx;
}
.btn-prev { background: #f0f0f0; color: #666; border: none; border-radius: 12rpx; padding: 20rpx 60rpx; font-size: 28rpx; }
.btn-next { background: #1a1a2e; color: #fff; border: none; border-radius: 12rpx; padding: 20rpx 60rpx; font-size: 28rpx; flex: 1; margin-left: 20rpx; }
.btn-submit { background: #e94560; color: #fff; border: none; border-radius: 12rpx; padding: 20rpx 60rpx; font-size: 28rpx; flex: 1; margin-left: 20rpx; }
</style>

