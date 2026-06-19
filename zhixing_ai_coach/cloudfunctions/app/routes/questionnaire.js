const { db } = require('../database')

exports.handle = async (event, userId) => {
  const method = event.httpMethod
  const memberId = event.path.split('/').pop()

  if (method === 'POST') {
    const body = typeof event.body === 'string' ? JSON.parse(event.body) : event.body
    const { member_id, raw_data } = body

    const { data: members } = await db.collection('members').where({ _id: member_id, userId }).get()
    if (!members.length) throw new Error('会员不存在')

    if (members[0].questionnaireId) throw new Error('问卷已提交')

    const result = await db.collection('questionnaires').add({
      data: { memberId: member_id, version: 'V3', rawData: JSON.stringify(raw_data), createdAt: db.serverDate() },
    })
    await db.collection('members').doc(member_id).update({
      data: { questionnaireId: result.id, status: 'questionnaire_done', updatedAt: db.serverDate() },
    })

    const { data: q } = await db.collection('questionnaires').doc(result.id).get()
    return { id: q._id, member_id: q.memberId, version: q.version, raw_data: JSON.parse(q.rawData), created_at: q.createdAt }
  }

  if (method === 'GET') {
    const { data: members } = await db.collection('members').where({ _id: memberId, userId }).get()
    if (!members.length) throw new Error('会员不存在')
    const qId = members[0].questionnaireId
    if (!qId) throw new Error('问卷尚未提交')

    const { data: q } = await db.collection('questionnaires').doc(qId).get()
    return { id: q._id, member_id: q.memberId, version: q.version, raw_data: JSON.parse(q.rawData), created_at: q.createdAt }
  }

  throw new Error('Method not allowed')
}
