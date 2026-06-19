const axios = require('axios')
const config = require('../config')

const PROMPT_TEMPLATE = `你是一位拥有8年一线私教经验的专业健身教练，名叫崔知行。

=== 用户信息 ===
姓名：{name}
年龄：{age}
性别：{gender}
训练经验：{experience}
每周频率：{frequency}次
训练目标：{goals}
目标细节：{goalDetails}
器械条件：{equipment}
工作压力：{workStress}
睡眠质量：{sleep}
运动损伤：{injury}
心理状态：{psychology}
会员类型：{memberType}
核心策略：{coreStrategy}

=== 输出要求 ===
请生成个性化训练方案（中文），包含：
1. 方案概览
2. 训练日安排
3. 每个训练日的详细计划（含呼吸、纠偏、发力感）
4. 渐进负荷计划（4周）
5. 备用方案（精简版+在家版）
6. 安全提示
`

async function generatePlan(memberInfo, classification) {
  const apiKey = config.AI_API_KEY
  if (!apiKey) {
    return generateFallbackPlan(memberInfo, classification)
  }

  const prompt = PROMPT_TEMPLATE
    .replace('{name}', memberInfo.name || '用户')
    .replace('{age}', memberInfo.age || '未知')
    .replace('{gender}', memberInfo.gender || '未知')
    .replace('{experience}', memberInfo.experience || '未知')
    .replace('{frequency}', String(memberInfo.frequency || 3))
    .replace('{goals}', memberInfo.goals || '增肌减脂')
    .replace('{goalDetails}', memberInfo.goalDetails || '')
    .replace('{equipment}', memberInfo.equipment || '健身房')
    .replace('{workStress}', memberInfo.workStress || '正常')
    .replace('{sleep}', memberInfo.sleep || '一般')
    .replace('{injury}', memberInfo.injury || '无')
    .replace('{psychology}', memberInfo.psychology || '正常')
    .replace('{memberType}', classification.memberType || 'TYPE_A')
    .replace('{coreStrategy}', classification.coreStrategy || '')

  try {
    const resp = await axios.post(config.AI_API_URL, {
      model: config.AI_MODEL,
      max_tokens: 8192,
      messages: [{ role: 'user', content: prompt }],
    }, {
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      timeout: 120000,
    })
    return resp.data.content[0].text
  } catch (e) {
    return `AI生成失败: ${e.message}\n\n请检查AI_API_KEY配置`
  }
}

function generateFallbackPlan(memberInfo, classification) {
  const name = memberInfo.name || '用户'
  const freq = memberInfo.frequency || 3
  return `## 训练方案（待AI生成）

### 方案概览
- **会员**: ${name}
- **会员类型**: ${classification.memberType || '待分类'}
- **训练频率**: 每周${freq}练
- **核心策略**: ${classification.coreStrategy || ''}

### 待AI生成内容
请配置 AI_API_KEY 后在CloudBase环境变量中设置。`
}

module.exports = { generatePlan }
