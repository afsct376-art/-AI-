from dataclasses import dataclass


MEMBER_TYPES = {
    "TYPE_A": "渐进入门型",
    "TYPE_B": "重建信心型",
    "TYPE_C": "进阶突破型",
    "TYPE_D": "精细优化型",
    "TYPE_E": "同行教练型",
    "TYPE_F": "老会员进阶型",
    "TYPE_G": "高压工作型",
}

TYPE_STRATEGIES = {
    "TYPE_A": {
        "split": "全身训练",
        "cycle": "CYCLE_SIMPLE",
        "voice": "VOICE_WARM",
        "core": "第一周不训练+兜底+PMS+粗粮",
        "modules": ["MOD_MENTAL"],
    },
    "TYPE_B": {
        "split": "全身训练",
        "cycle": "CYCLE_2DAY",
        "voice": "VOICE_AFFIRMATIVE",
        "core": "继续旅程叙事+双路径+夜班方案",
        "modules": ["MOD_MENTAL", "MOD_NOBLAME"],
    },
    "TYPE_C": {
        "split": "推拉腿分化",
        "cycle": "CYCLE_2DAY",
        "voice": "VOICE_DATA",
        "core": "DUP+RPE+器械突破+季度全休",
        "modules": ["MOD_DECISION_TREE", "MOD_MENTAL"],
    },
    "TYPE_D": {
        "split": "上下肢分化",
        "cycle": "CYCLE_2DAY",
        "voice": "VOICE_ELITE",
        "core": "左右平衡+赛前减量+专项营养",
        "modules": ["MOD_DECISION_TREE", "MOD_MENTAL"],
    },
    "TYPE_E": {
        "split": "DUP",
        "cycle": "CYCLE_2DAY",
        "voice": "VOICE_PEER",
        "core": "专业框架+术语速查+示例FAQ",
        "modules": ["MOD_DECISION_TREE", "MOD_MENTAL"],
    },
    "TYPE_F": {
        "split": "上下肢分化",
        "cycle": "CYCLE_2DAY",
        "voice": "VOICE_PRO",
        "core": "线性周期+上下肢分化",
        "modules": ["MOD_MENTAL"],
    },
    "TYPE_G": {
        "split": "两分化（上肢日/下肢日轮换）",
        "cycle": "CYCLE_ABC",
        "voice": "VOICE_GENTLE",
        "core": "ABC三方案+成绩单封面+完成比完美",
        "modules": ["MOD_ABC", "MOD_SCORECARD", "MOD_KNOWN_ACTIONS", "MOD_NOBLAME", "MOD_MENTAL"],
    },
}


@dataclass
class MemberProfile:
    experience: int       # D1: 0-5
    frequency: int        # D2: 每周天数
    equipment: int        # D3: 0-4
    injury: int           # D4: 0-6
    lifestyle: int        # D5: 0-5
    psychology: int       # D6: 0-5
    training_history: int # D7: 0-3
    special_needs: int    # D8: 0-2
    female_specific: int  # D9: 0-3
    motivation: int       # D10: 0-3
    is_restart: bool = False
    is_coach: bool = False
    gender: str = "male"
    night_shift_days: int = 0
    commute_minutes: int = 0
    peak_hours: list[str] | None = None
    max_streak_months: int = 0
    expect_result_weeks: int = 12


def classify(profile: MemberProfile) -> dict:
    result = {
        "member_type": "",
        "split": "",
        "cycle": "",
        "voice": "",
        "core_strategy": "",
        "enabled_modules": [],
        "sub_strategies": {},
    }

    # Step 1: Core type classification
    if profile.psychology == 5:
        result["member_type"] = "TYPE_G"
    elif profile.experience == 0:
        result["member_type"] = "TYPE_A"
    elif profile.is_restart and profile.psychology == 2:
        result["member_type"] = "TYPE_B"
    elif profile.is_coach:
        result["member_type"] = "TYPE_E"
    elif profile.experience >= 3:
        result["member_type"] = "TYPE_D"
    elif profile.experience == 2:
        result["member_type"] = "TYPE_C"
    else:
        result["member_type"] = "TYPE_F"

    mt = result["member_type"]
    strategy = TYPE_STRATEGIES[mt]
    result["split"] = strategy["split"]
    result["cycle"] = strategy["cycle"]
    result["voice"] = strategy["voice"]
    result["core_strategy"] = strategy["core"]
    result["enabled_modules"] = list(strategy["modules"])

    # Step 2: Lifestyle sub-strategies
    if profile.lifestyle >= 3 or profile.night_shift_days >= 4:
        result["sub_strategies"]["night_shift"] = "双路径方案——夜班日补觉+白班日训练"
        result["sub_strategies"]["frequency_model"] = "不固定周频——按班型动态调整"
        if "MOD_NIGHTSHIFT" not in result["enabled_modules"]:
            result["enabled_modules"].append("MOD_NIGHTSHIFT")

    if profile.commute_minutes >= 60:
        result["sub_strategies"]["time_budget_adjustment"] = "可用训练时间减15分钟"
        if "MOD_COMMUTE" not in result["enabled_modules"]:
            result["enabled_modules"].append("MOD_COMMUTE")

    # Step 3: Peak hours → equipment substitution
    if profile.peak_hours and ("17-20" in str(profile.peak_hours) or "17:00-20:00" in str(profile.peak_hours)):
        if "MOD_HOUR" not in result["enabled_modules"]:
            result["enabled_modules"].append("MOD_HOUR")

    # Step 4: Female-specific
    if profile.gender == "female" and profile.female_specific > 0:
        if profile.female_specific <= 1:
            result["sub_strategies"]["pms"] = "经期前2天降级为精简版，之后恢复正常"
        elif profile.female_specific <= 2:
            result["sub_strategies"]["pms"] = "经期前3天切换为在家版，后2天过渡"
        else:
            result["sub_strategies"]["pms"] = "经期全程切换为精简方案，经期后恢复正常"
        if "MOD_PMS" not in result["enabled_modules"]:
            result["enabled_modules"].append("MOD_PMS")

    # Step 5: Motivation → expectation management
    if profile.max_streak_months < 1:
        result["sub_strategies"]["expectation"] = "短期反馈密集——每周一次进度确认；方案极简——先养成习惯再优化"
    if profile.expect_result_weeks <= 4:
        result["sub_strategies"]["visual_feedback"] = "可视化设计——每周拍照+围度对比"
        if "MOD_EXPECTATION" not in result["enabled_modules"]:
            result["enabled_modules"].append("MOD_EXPECTATION")

    # Step 6: Private coach transition
    if profile.training_history == 2:
        if "MOD_KNOWN_ACTIONS" not in result["enabled_modules"]:
            result["enabled_modules"].append("MOD_KNOWN_ACTIONS")

    return result
