import json
import httpx
from typing import Any

from app.config import settings

PROMPT_TEMPLATE = """
你是一位拥有8年一线私教经验的专业健身教练，名叫崔知行。你的专长是：

1. 为不同目标（增肌/减脂/塑形/力量提升/综合体能）设计个性化训练方案
2. 善于在有限器械条件下做最优动作选择
3. 关注动作安全和渐进负荷，尤其擅长膝关节/肩关节/下背的风险管理
4. 注意训练方案的实际可执行性——用户能否长期坚持比理论完美更重要
5. 善于帮用户建立自主训练能力——不只给方案，还要教会用户自己做决策
6. 能识别用户的心理状态并匹配对应的沟通策略

=== 用户信息 ===

【基本信息】
- 姓名：{name}
- 年龄：{age}
- 性别：{gender}
- 身高：{height_cm}
- 体重：{weight_kg}
- 训练经验：{experience_level}
- 当前训练频率：每周{frequency}次
- 过往教练经历：{coach_history}

【目标设定】
- 主要目标：{goals}
- 目标细节：{goal_details}
- 期望达成时间：{expect_time}

【器械条件】
- 训练地点：{venue}
- 可用器械：{equipment_list}
- 器械上限：{equipment_limits}

【时间与偏好】
- 每次可用时间：{session_time}分钟
- 每周可用天数：{weekly_days}天
- 偏好时间段：{preferred_time}
- 不喜欢动作：{disliked_exercises}

【生活方式】
- 工作压力：{work_stress}
- 应酬频率：{social_frequency}
- 睡眠质量：{sleep_quality}
- 通勤时间：{commute}分钟
- 夜班情况：{night_shift}

【运动损伤】
- 伤病史：{injury_history}
- 当前不适：{current_pain}
- 关节问题：{joint_issues}

【心理状态】
- 当前状态：{psychology_state}
- 特殊关注：{special_attention}

【会员分类结果】
- 会员类型：{member_type}
- 核心策略：{core_strategy}
- 启用模块：{enabled_modules}

=== 输出要求 ===

请按以下结构生成个性化训练方案（使用中文）：

## 方案概览
## 训练日安排
## 每个训练日的详细计划
## 有氧安排
## 热身与冷身
## 渐进负荷计划（4周）
## 备用方案（精简版+在家版）
## 营养建议
## 安全提示
## 下周调整指南
## 4周后复盘节点

注意：
- 每个动作要点必须包含呼吸、纠偏、发力感三要素
- 发力感用生活化比喻而非解剖学术语
- 如果用户工作压力大→提供ABC三方案（完整版/精简版/在家版）
- 如果用户有关节不适→增加关节保护专项模块
- 如果用户是同行教练→增加术语速查表
"""


async def generate_plan(member_info: dict, classification: dict) -> str:
    api_key = settings.AI_API_KEY
    if not api_key:
        return _generate_fallback_plan(member_info, classification)

    prompt = PROMPT_TEMPLATE.format(
        name=member_info.get("name", "用户"),
        age=member_info.get("age", "未知"),
        gender=member_info.get("gender", "未知"),
        height_cm=member_info.get("height_cm", "未知"),
        weight_kg=member_info.get("weight_kg", "未知"),
        experience_level=member_info.get("experience_level", "未知"),
        frequency=member_info.get("frequency", 3),
        coach_history=member_info.get("coach_history", "无"),
        goals=member_info.get("goals", "增肌减脂"),
        goal_details=member_info.get("goal_details", ""),
        expect_time=member_info.get("expect_time", "3个月"),
        venue=member_info.get("venue", "商业健身房"),
        equipment_list=member_info.get("equipment_list", "基本器械"),
        equipment_limits=member_info.get("equipment_limits", "无"),
        session_time=member_info.get("session_time", 60),
        weekly_days=member_info.get("weekly_days", 3),
        preferred_time=member_info.get("preferred_time", "晚上"),
        disliked_exercises=member_info.get("disliked_exercises", "无"),
        work_stress=member_info.get("work_stress", "正常"),
        social_frequency=member_info.get("social_frequency", "无"),
        sleep_quality=member_info.get("sleep_quality", "一般"),
        commute=member_info.get("commute", 30),
        night_shift=member_info.get("night_shift", "无"),
        injury_history=member_info.get("injury_history", "无"),
        current_pain=member_info.get("current_pain", "无"),
        joint_issues=member_info.get("joint_issues", "无"),
        psychology_state=member_info.get("psychology_state", "正常"),
        special_attention=member_info.get("special_attention", "无"),
        member_type=classification.get("member_type", "TYPE_A"),
        core_strategy=classification.get("core_strategy", ""),
        enabled_modules=", ".join(classification.get("enabled_modules", [])),
    )

    try:
        async with httpx.AsyncClient(timeout=120.0) as client:
            resp = await client.post(
                settings.AI_API_URL,
                headers={
                    "x-api-key": api_key,
                    "anthropic-version": "2023-06-01",
                    "content-type": "application/json",
                },
                json={
                    "model": settings.AI_MODEL,
                    "max_tokens": 8192,
                    "messages": [{"role": "user", "content": prompt}],
                },
            )
            resp.raise_for_status()
            data = resp.json()
            return data["content"][0]["text"]
    except Exception as e:
        return f"AI生成失败: {e}\n\n请使用备用Prompt手动生成。"


def _generate_fallback_plan(member_info: dict, classification: dict) -> str:
    """When no AI API key is configured, generate a template plan."""
    mt = classification.get("member_type", "TYPE_A")
    name = member_info.get("name", "用户")
    freq = member_info.get("frequency", 3)

    type_names = {
        "TYPE_A": "渐进入门型",
        "TYPE_G": "高压工作型",
    }
    type_name = type_names.get(mt, mt)

    return f"""## 训练方案（模板 — 待AI生成）

### 方案概览
- **会员**: {name}
- **会员类型**: {type_name}
- **训练频率**: 每周{freq}练
- **核心策略**: {classification.get("core_strategy", "")}

### 待AI生成内容
请配置 AI_API_KEY 后重新生成完整方案，包含：
1. 详细训练日安排
2. 每个动作的呼吸/纠偏/发力感
3. 渐进负荷计划
4. 备用方案
5. 营养建议
6. 安全提示

### 已启用的策略模块
{chr(10).join(f"- {m}" for m in classification.get("enabled_modules", []))}
"""
