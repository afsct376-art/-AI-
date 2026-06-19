const { db } = require('../database')
const { classify, buildProfile } = require('../services/decisionTree')
const { generatePlan } = require('../services/aiEngine')

exports.handle = async (event, userId) => {
  const method = event.httpMethod
  const body = typeof event.body === 'string' ? JSON.parse(event.body) : event.body
  const memberId = method === 'GET' ? event.path.split('/').pop() : body.member_id

  const { data: members } = await db.collection('members').where({ _id: memberId, userId }).get()
  if (!members.length) throw new Error('会员不存在')
  const member = members[0]

  // GET: return plan
  if (method === 'GET') {
    if (!member.planId) throw new Error('方案尚未生成')
    const { data: plans } = await db.collection('plans').doc(member.planId).get()
    return plans
  }

  // POST: generate plan
  if (method === 'POST') {
    if (!member.questionnaireId) throw new Error('请先提交问卷')

    const qId = member.questionnaireId
    const { data: questionnaires } = await db.collection('questionnaires').doc(qId).get()
    const qData = JSON.parse(questionnaires.rawData)

    const profile = buildProfile(qData)
    const classification = classify(profile)

    const memberInfo = {
      name: member.name,
      frequency: profile.frequency,
      goals: qData.Q9 || '',
      goalDetails: qData.Q10 || '',
    }

    const planContent = await generatePlan(memberInfo, classification)

    const planData = {
      memberId, content: planContent, planType: classification.memberType,
      splitType: classification.split, weeklyFrequency: profile.frequency,
      status: 'ready', coachNotes: body.coach_notes || '', createdAt: db.serverDate(), updatedAt: db.serverDate(),
    }

    let planId = member.planId
    if (planId) {
      await db.collection('plans').doc(planId).update({ data: planData })
    } else {
      const result = await db.collection('plans').add({ data: planData })
      planId = result.id
    }

    await db.collection('members').doc(memberId).update({
      data: {
        planId, memberType: classification.memberType, status: 'completed',
        dimExperience: profile.experience, dimFrequency: profile.frequency,
        dimEquipment: profile.equipment, dimInjury: profile.injury,
        dimLifestyle: profile.lifestyle, dimPsychology: profile.psychology,
        dimTrainingHistory: profile.trainingHistory, dimSpecialNeeds: profile.specialNeeds,
        dimFemaleSpecific: profile.femaleSpecific, dimMotivation: profile.motivation,
        updatedAt: db.serverDate(),
      },
    })

    const { data: plan } = await db.collection('plans').doc(planId).get()
    return plan
  }

  throw new Error('Method not allowed')
}
