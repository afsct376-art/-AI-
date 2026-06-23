# OpenCode 高级特性练习指南

## 专业术语解释

### 基础概念

| 术语 | 说明 |
|------|------|
| **LLM** | 大语言模型，即 AI 大脑（如 DeepSeek、GPT、Gemini） |
| **Provider** | 模型提供商，提供 AI 模型 API 的公司（如 OpenAI、Google、Anthropic） |
| **Agent** | AI 代理，能自主思考、调用工具、读写文件来完成任务 |
| **Primary Agent** | 主代理，你直接对话的那个 AI，可以按 Tab 切换（Build / Plan） |
| **Subagent** | 子代理，专门干一件事的小助手，用 `@名字` 调用 |
| **Prompt** | 提示词，你发给 AI 的指令 |
| **Context** | 上下文，AI 当前对话能看到的所有信息（会消耗 token） |
| **Token** | AI 计费/计量的基本单位，≈ 0.75 个中文字或 4 个英文字母 |
| **YAML** | 一种配置文件格式，用 `---` 包裹的键值对，比 JSON 更易读 |

### 免费模型

通过 `/connect` → **OpenCode Zen** 获取免费额度，当前免费模型：
- `deepseek-v4-flash-free` — DeepSeek 最新的免费模型，速度快
- `big-pickle` — 代号"大 pickle"，来自新供应商的测试模型
- `mimo-v2.5-free` — Minimax 公司的免费模型
- `nemotron-3-ultra-free` — NVIDIA 的免费模型

或 `/connect` → **Google Gemini**（250次/天，免费，无需绑信用卡）。

---

## 1. Agent Skills（技能）

Skill = **可复用的指令模板文件**

> 就像给 AI 一本"操作手册"，告诉它特定场景下该怎么做。比如写一个 `code-review` skill，以后让它审查代码就会遵循你的规则。

### 创建步骤
1. 新建目录 `.opencode/skills/<skill-name>/SKILL.md`
2. 写入 YAML 头部 + 指令内容
3. 启动 `opencode`，agent 会自动发现

### 格式
```yaml
---
name: my-skill
description: 简短描述
---
你的指令内容...
```

### 练习项目
- 做一个 `code-review` skill，自动审查代码风格
- 做一个 `git-helper` skill，自动生成 commit message

---

## 2. MCP Server（模型上下文协议服务器）

MCP = **给 AI 接入外部工具的"万能插头"**

> 想象 AI 本来只有"手"（读写文件）和"眼睛"（搜索代码）。MCP 就像给它装上了各种"外挂装备"——可以操作浏览器、查数据库、发 GitHub PR、搜网页。每个 MCP Server 就是一个工具包。

### 配置方式
编辑 `opencode.json` 添加：
```json
"mcp": {
  "server-name": {
    "type": "local",
    "command": ["npx", "-y", "mcp-package-name"]
  }
}
```

### 练习项目
- `@modelcontextprotocol/server-everything` — MCP 官方测试工具包，用来验证 MCP 是否配置成功
- `@playwright/mcp` — 让 AI 能打开浏览器、点击按钮、截图，用来做自动化测试
- `@modelcontextprotocol/server-github` — 让 AI 能直接操作 GitHub（创建 issue、PR、搜索代码）

### 使用
在 prompt 中加：`use <server-name> tool to ...`

---

## 3. Subagent（子代理）

Subagent = **专门干一件事的"小助手"**

> 相当于你有一个"主力 AI"，然后手下还有几个"实习生"。主力 AI 忙不过来时可以把活派给实习生干，你也可以直接喊实习生干活。

### 内置免费 subagents
- `@explore` — 快速搜索代码（只会读不会改）
- `@scout` — 查外部文档和 npm 包源码
- `@general` — 万能的，啥都能干（但不能写 todo）

### 自定义 subagent
在 `opencode.json` 中配置：
```json
"agent": {
  "test-writer": {
    "description": "专门写测试的助手",
    "mode": "subagent",
    "permission": { "edit": "allow", "bash": "deny" },
    "prompt": "你是一个测试工程师，用 vitest 写单元测试"
  }
}
```

字段说明：
| 字段 | 含义 |
|------|------|
| `description` | 描述，告诉 AI 这个子代理是干嘛的 |
| `mode` | 模式，`subagent` 表示是子代理，`primary` 是主代理 |
| `permission` | 权限，`edit: allow` 允许改文件，`bash: deny` 禁止执行命令 |
| `prompt` | 提示词，给这个子代理的专属指令 |

### 调用方式
在对话中输入：`@test-writer 给 src/index.ts 写测试`

### 练习项目
创建 `test-writer`、`docs-writer`、`db-helper` 三个自定义 subagent

---

## 4. Agent 编排（组合使用）

> 把上面的技能全部串起来，让多个 AI 角色协同工作

### Workflow
1. `Tab` 切换到 Plan 模式做分析（不写代码，只出方案）
2. `Tab` 切回 Build 模式执行（动手写代码）
3. 用 `@explore` 搜索 → `@general` 修改 → Build 汇总

### 全链路练习
```
@explore 找到所有 API 路由文件
→ plan agent 分析如何加鉴权
→ build agent 执行修改
→ MCP playwright 测试结果
```

---

## 技术名词速查表

| 英文术语 | 中文 | 大白话解释 |
|----------|------|-----------|
| Agent | 代理 | AI 替你干活的"数字员工" |
| Subagent | 子代理 | 专门干某一类活的"专科助手" |
| Skill | 技能 | 给 AI 的"操作手册" |
| MCP | 模型上下文协议 | 给 AI 装"外挂"的标准化接口 |
| Plugin | 插件 | 用 TypeScript 写的自定义扩展 |
| LSP | 语言服务器协议 | 让 AI 能看懂代码语法错误 |
| Provider | 供应商 | 提供 AI 模型的公司 |
| Token | 令牌 | AI 处理信息的最小单位 |
| Context | 上下文窗口 | AI 一次能记住的最大信息量 |
| Prompt | 提示词 | 你发给 AI 的指令文本 |
| Permission | 权限 | 允许或禁止 AI 做某事（读写文件、执行命令等） |
| YAML | 一种配置格式 | 用缩进和冒号写配置，比 JSON 简洁 |
