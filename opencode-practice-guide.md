# OpenCode 高级特性练习指南

## 免费模型

通过 `/connect` → **OpenCode Zen** 获取免费额度，当前免费模型：
- `deepseek-v4-flash-free`
- `big-pickle`
- `mimo-v2.5-free`
- `nemotron-3-ultra-free`

或 `/connect` → **Google Gemini**（250次/天，免费）。

---

## 1. Agent Skills

Skill = 可复用的指令模板文件

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

## 2. MCP Server

MCP = 给 AI 接入外部工具

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
- 安装 `@modelcontextprotocol/server-everything` 测试基础功能
- 安装 `@playwright/mcp`，让 AI 自动打开浏览器测试网页
- 安装 `@modelcontextprotocol/server-github` 操作 GitHub

### 使用
在 prompt 中加：`use <server-name> tool to ...`

---

## 3. Subagent

Subagent = 专门干一件事的"小助手"

### 内置免费 subagents
- `@explore` — 快速搜索代码（只读）
- `@scout` — 查外部文档和依赖
- `@general` — 多步骤任务

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

### 调用方式
在对话中输入：`@test-writer 给 src/index.ts 写测试`

### 练习项目
创建 `test-writer`、`docs-writer`、`db-helper` 三个自定义 subagent

---

## 4. Agent 编排（组合使用）

### Workflow
1. `Tab` 切换到 Plan 模式做分析
2. `Tab` 切回 Build 模式执行
3. 用 `@explore` 搜索 → `@general` 修改 → Build 汇总

### 全链路练习
```
@explore 找到所有 API 路由文件
→ plan agent 分析如何加鉴权
→ build agent 执行修改
→ MCP playwright 测试结果
```
