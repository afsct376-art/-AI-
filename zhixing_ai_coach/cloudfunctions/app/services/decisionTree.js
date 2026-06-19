const MEMBER_TYPES = [
  { id: 'TYPE_A', name: '渐进入门型', core: '第一周不训练+兜底+PMS+粗粮', split: '全身训练', voice: 'VOICE_WARM', modules: ['MOD_MENTAL'] },
  { id: 'TYPE_B', name: '重建信心型', core: '继续旅程叙事+双路径+夜班方案', split: '全身训练', voice: 'VOICE_AFFIRMATIVE', modules: ['MOD_MENTAL', 'MOD_NOBLAME'] },
  { id: 'TYPE_C', name: '进阶突破型', core: 'DUP+RPE+器械突破+季度全休', split: '推拉腿分化', voice: 'VOICE_DATA', modules: ['MOD_DECISION_TREE', 'MOD_MENTAL'] },
  { id: 'TYPE_D', name: '精细优化型', core: '左右平衡+赛前减量+专项营养', split: '上下肢分化', voice: 'VOICE_ELITE', modules: ['MOD_DECISION_TREE', 'MOD_MENTAL'] },
  { id: 'TYPE_E', name: '同行教练型', core: '专业框架+术语速查+示例FAQ', split: 'DUP', voice: 'VOICE_PEER', modules: ['MOD_DECISION_TREE', 'MOD_MENTAL'] },
  { id: 'TYPE_F', name: '老会员进阶型', core: '线性周期+上下肢分化', split: '上下肢分化', voice: 'VOICE_PRO', modules: ['MOD_MENTAL'] },
  { id: 'TYPE_G', name: '高压工作型', core: 'ABC三方案+成绩单封面+完成比完美', split: '两分化（上肢日/下肢日轮换）', voice: 'VOICE_GENTLE', modules: ['MOD_ABC', 'MOD_SCORECARD', 'MOD_KNOWN_ACTIONS', 'MOD_NOBLAME', 'MOD_MENTAL'] },
]

function classify(profile) {
  const result = { memberType: '', split: '', voice: '', coreStrategy: '', enabledModules: [], subStrategies: {} }

  // Step 1: Core type
  if (profile.psychology === 5) result.memberType = 'TYPE_G'
  else if (profile.experience === 0) result.memberType = 'TYPE_A'
  else if (profile.isRestart && profile.psychology === 2) result.memberType = 'TYPE_B'
  else if (profile.isCoach) result.memberType = 'TYPE_E'
  else if (profile.experience >= 3) result.memberType = 'TYPE_D'
  else if (profile.experience === 2) result.memberType = 'TYPE_C'
  else result.memberType = 'TYPE_F'

  const strategy = MEMBER_TYPES.find(t => t.id === result.memberType)
  if (strategy) {
    result.split = strategy.split
    result.voice = strategy.voice
    result.coreStrategy = strategy.core
    result.enabledModules = [...strategy.modules]
  }

  // Step 2: Lifestyle
  if (profile.nightShiftDays >= 4) {
    result.subStrategies.nightShift = '双路径方案——夜班日补觉+白班日训练'
    if (!result.enabledModules.includes('MOD_NIGHTSHIFT')) result.enabledModules.push('MOD_NIGHTSHIFT')
  }
  if (profile.commuteMinutes >= 60) {
    result.subStrategies.timeBudget = '可用训练时间减15分钟'
    if (!result.enabledModules.includes('MOD_COMMUTE')) result.enabledModules.push('MOD_COMMUTE')
  }

  // Step 3: Female
  if (profile.gender === 'female' && profile.femaleSpecific > 0) {
    const map = { 1: '经期前2天降级为精简版', 2: '经期前3天切换为在家版', 3: '经期全程切换为精简方案' }
    result.subStrategies.pms = map[profile.femaleSpecific] || '经期调整'
    if (!result.enabledModules.includes('MOD_PMS')) result.enabledModules.push('MOD_PMS')
  }

  // Step 4: Motivation
  if (profile.maxStreakMonths < 1) {
    result.subStrategies.expectation = '短期反馈密集——每周一次进度确认'
  }

  return result
}

function buildProfile(q) {
  const expMap = { '完全新手': 0, '初学者': 1, '曾经练过': 1, '中级': 2, '高级': 3, '同行教练': 4 }
  let exp = 0
  const expRaw = q.Q6 || ''
  for (const [k, v] of Object.entries(expMap)) { if (expRaw.includes(k)) { exp = v; break } }

  const freq = parseInt((q.Q7 || '3').replace('次', '')) || 3

  const venue = q.Q11 || ''
  let equip = 2
  if (venue.includes('徒手')) equip = 0
  else if (venue.includes('家庭')) equip = 1
  else if (venue.includes('商业') || venue.includes('铁馆')) equip = 4

  const joints = q['Q13-extra'] || ''
  const injuryMap = { '膝关节': 1, '下背': 2, '肩关节': 3, '腰椎': 2, '术后': 4, '踝关节': 5 }
  let injury = 0
  for (const [k, v] of Object.entries(injuryMap)) { if (joints.includes(k)) injury = Math.max(injury, v) }

  const stress = q.Q19c || '正常'
  let lifestyle = 0
  if (stress.includes('极大') || stress.includes('加班')) lifestyle += 2
  else if (stress.includes('偏大')) lifestyle += 1
  const social = q.Q20 || '几乎没有'
  if (social.includes('2次')) lifestyle += 1
  else if (social.includes('3次')) lifestyle += 2

  const psychMap = {
    '很想开始但一直没动起来': 0,
    '之前练得好，停了一段时间，不知道怎么重新开始': 2,
    '有基础，但感觉遇到了瓶颈': 1,
    '工作太忙/压力大，练了又断，有点内疚': 5,
    '一直在规律训练，想要继续进阶': 3,
    '健身教练，来找专业交流': 4,
  }
  const psychology = psychMap[q.Q19f] || 0

  const coachQ = q.Q8a || ''
  const trainingHistory = coachQ.includes('跟过') || coachQ.includes('还在跟') ? 2 : coachQ.includes('没跟过') ? 0 : 1

  const gender = q.Q3 || '男'
  const impact = q.Q25a || '不适用'
  let female = 0
  if (gender === '女') {
    if (impact.includes('前1-2天')) female = 1
    else if (impact.includes('前3天')) female = 2
    else if (impact.includes('完全不')) female = 3
  }

  const streak = q.Q10a || ''
  let maxStreak = 0
  if (streak.includes('1-3个月')) maxStreak = 2
  else if (streak.includes('3-6个月')) maxStreak = 4
  else if (streak.includes('6个月以上')) maxStreak = 6

  const nightRaw = q['Q19d-extra'] || '无夜班'
  let nightDays = 0
  if (nightRaw.includes('4-7天')) nightDays = 4
  else if (nightRaw.includes('8-15天')) nightDays = 8
  else if (nightRaw.includes('15天以上')) nightDays = 15

  const commuteRaw = q.Q19e || '30分钟内'
  let commute = 0
  if (commuteRaw.includes('30-60')) commute = 30
  else if (commuteRaw.includes('60-90')) commute = 60
  else if (commuteRaw.includes('90')) commute = 90

  return {
    experience: exp, frequency: freq, equipment: equip, injury, lifestyle, psychology,
    trainingHistory, femaleSpecific: female, specialNeeds: 0, motivation: 0,
    isRestart: expRaw.includes('重新开始') || (q.Q19f || '').includes('停了一段时间'),
    isCoach: (q.Q19f || '').includes('健身教练'), gender, nightShiftDays: nightDays,
    commuteMinutes: commute, maxStreakMonths: maxStreak, expectResultWeeks: 12,
  }
}

module.exports = { classify, buildProfile, MEMBER_TYPES }
