const { db } = require('../database')

exports.handle = async (event, userId) => {
  const method = event.httpMethod

  if (method === 'POST') {
    const result = await db.collection('members').add({
      data: { userId, name: event.name || '用户', memberType: '', status: 'pending', dimExperience: 0, dimFrequency: 0, dimEquipment: 0, dimInjury: 0, dimLifestyle: 0, dimPsychology: 0, dimTrainingHistory: 0, dimSpecialNeeds: 0, dimFemaleSpecific: 0, dimMotivation: 0, coachNotes: '', createdAt: db.serverDate(), updatedAt: db.serverDate() },
    })
    const { data } = await db.collection('members').doc(result.id).get()
    return data
  }

  if (method === 'GET') {
    const { data } = await db.collection('members').where({ userId }).orderBy('createdAt', 'desc').get()
    const list = data.map(m => {
      return { ...m, questionnaireSubmitted: !!(m.questionnaireId), planReady: m.status === 'completed', hasPlan: !!(m.planId) }
    })
    return { members: list, total: list.length }
  }

  throw new Error('Method not allowed')
}
