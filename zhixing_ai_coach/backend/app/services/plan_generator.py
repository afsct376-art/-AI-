import json
from typing import Any

from app.services.decision_tree import MemberProfile, classify
from app.services.ai_engine import generate_plan


def build_member_profile(questionnaire_data: dict[str, Any]) -> MemberProfile:
    """Convert raw questionnaire response into a MemberProfile for the decision tree."""
    q = questionnaire_data

    exp_map = {"完全新手": 0, "初学者": 1, "曾经练过": 1, "中级": 2, "高级": 3, "同行教练": 4}
    exp_raw = q.get("Q6", "")
    exp = 0
    for k, v in exp_map.items():
        if k in exp_raw:
            exp = v
            break

    freq = int(q.get("Q7", "3").replace("次", "")) if q.get("Q7") else 3

    equip = 2
    venue = q.get("Q11", "")
    if "徒手" in venue:
        equip = 0
    elif "家庭" in venue and "部分器械" in venue:
        equip = 1
    elif "商业" in venue or "铁馆" in venue:
        equip = 4

    injury_answers = {
        "膝关节": 1, "下背": 2, "肩关节": 3,
        "腰椎": 2, "术后": 4, "踝关节": 5,
    }
    injury = 0
    joint_str = q.get("Q13-extra", "")
    for k, v in injury_answers.items():
        if k in joint_str:
            injury = max(injury, v)

    lifestyle = 0
    work_stress = q.get("Q19c", "正常")
    if "极大" in work_stress or "加班" in work_stress:
        lifestyle += 2
    elif "偏大" in work_stress:
        lifestyle += 1
    social = q.get("Q20", "几乎没有")
    if "2次" in social:
        lifestyle += 1
    elif "3次" in social:
        lifestyle += 2

    psych_map = {
        "很想开始但一直没动起来": 0,
        "之前练得好，停了一段时间，不知道怎么重新开始": 2,
        "有基础，但感觉遇到了瓶颈": 1,
        "工作太忙/压力大，练了又断，有点内疚": 5,
        "一直在规律训练，想要继续进阶": 3,
        "健身教练，来找专业交流": 4,
    }
    psych_raw = q.get("Q19f", "")
    psychology = psych_map.get(psych_raw, 0)

    coach_exp = q.get("Q8a", "")
    training_history = 0
    if "跟过" in coach_exp or "还在跟" in coach_exp:
        training_history = 2
    elif "没跟过" not in coach_exp:
        training_history = 1

    special_needs = 0
    goal = q.get("Q9", "")
    goal_detail = q.get("Q10", "")
    if "备赛" in goal or "备赛" in goal_detail:
        special_needs = 1
    if "学会自己做" in goal or "学会自己做" in goal_detail:
        special_needs = max(special_needs, 1)

    female = 0
    gender = q.get("Q3", "男")
    impact_str = q.get("Q25a", "不适用")
    if gender == "女":
        if "前1-2天" in impact_str:
            female = 1
        elif "前3天" in impact_str:
            female = 2
        elif "完全不" in impact_str:
            female = 3
        elif "不影响" in impact_str:
            female = 0

    streak = q.get("Q10a", "")
    max_streak = 0
    if "从没超过2周" in streak:
        max_streak = 0
    elif "2-4周" in streak:
        max_streak = 0
    elif "1-3个月" in streak:
        max_streak = 2
    elif "3-6个月" in streak:
        max_streak = 4
    elif "6个月以上" in streak:
        max_streak = 6

    expect = q.get("Q10b", "")
    expect_weeks = 12
    if "2周内" in expect:
        expect_weeks = 2
    elif "1个月内" in expect:
        expect_weeks = 4
    elif "2-3个月" in expect:
        expect_weeks = 10

    motivation = 1 if max_streak < 1 else (2 if expect_weeks <= 4 else 0)

    night_raw = q.get("Q19d-extra", "无夜班")
    night_days = 0
    if "4-7天" in night_raw:
        night_days = 4
    elif "8-15天" in night_raw:
        night_days = 8
    elif "15天以上" in night_raw:
        night_days = 15

    commute_raw = q.get("Q19e", "30分钟内")
    commute = 0
    if "30-60" in commute_raw:
        commute = 30
    elif "60-90" in commute_raw:
        commute = 60
    elif "90" in commute_raw:
        commute = 90

    peak = q.get("Q12a", "")

    is_coach = "健身教练" in q.get("Q19f", "")
    is_restart = "重新开始" in q.get("Q6", "") or "停了一段时间" in q.get("Q19f", "")

    return MemberProfile(
        experience=exp,
        frequency=freq,
        equipment=equip,
        injury=injury,
        lifestyle=lifestyle,
        psychology=psychology,
        training_history=training_history,
        special_needs=special_needs,
        female_specific=female,
        motivation=motivation,
        is_restart=is_restart,
        is_coach=is_coach,
        gender=gender,
        night_shift_days=night_days,
        commute_minutes=commute,
        peak_hours=[peak] if peak else None,
        max_streak_months=max_streak,
        expect_result_weeks=expect_weeks,
    )


async def create_plan_for_member(member_info: dict, questionnaire_data: dict[str, Any]) -> str:
    profile = build_member_profile(questionnaire_data)
    classification = classify(profile)
    plan_content = await generate_plan(member_info, classification)
    return plan_content
