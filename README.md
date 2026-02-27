# Claude for WPS PowerPoint

WPS Office 演示文稿 AI 助手——通过自然语言对话创建和编辑幻灯片，由 Claude API 驱动。

![License](https://img.shields.io/badge/license-MIT-blue)
![Version](https://img.shields.io/badge/version-2.0.0--dev-orange)
![Platform](https://img.shields.io/badge/platform-WPS%20Office-red)

## 功能特性

- **自然语言创建 PPT**：描述需求即可自动生成多页演示文稿，支持封面/目录/正文/结尾
- **专业视觉设计**：4 套配色方案（科技蓝/商务橙/自然绿/高端紫），自动选择最佳配色
- **逐页创建动画**：生成幻灯片时逐页翻页展示，直观看到 AI 的工作过程
- **三种交互模式**：Agent（自动执行）/ Plan（步骤规划）/ Ask（只读分析）
- **布局美化**：一键优化对齐、配色、字体、间距，提升视觉品质
- **流式响应**：SSE 流式输出 + Markdown 渲染 + 代码块语法高亮
- **代码执行桥**：生成的 WPP JS 代码可一键执行，支持结果回传和错误修复
- **会话历史**：自动保存对话记录，支持多会话切换和恢复
- **8 个快捷指令**：生成幻灯片、美化布局、添加备注、提取文本等
- **连接器系统**：通过 MCP 协议接入外部数据源（网络搜索、企业知识库）
- **模型选择**：Sonnet 4.6 / Opus 4.6 / Haiku 4.5

## 架构概览

```
┌─────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  WPS PPT    │◄───►│  Proxy Server    │◄───►│  React TaskPane │
│  Plugin Host│     │  (Express :3002) │     │  (Vite :5174)   │
│  main.js    │     │  proxy-server.js │     │  src/App.tsx     │
└─────────────┘     └──────────────────┘     └─────────────────┘
       │                     │                        │
       │ WPS WPP API         │ Claude CLI (SSE)       │ 用户对话
       │ Slides/Shapes       │ Skill/Mode 匹配       │ Markdown 渲染
       │ 代码执行+动画       │ Session 持久化         │ 代码高亮+执行
       │                     │ MCP Connectors         │ 模式切换
```

## 快速开始

### 前置条件

- **Node.js** >= 18
- **Claude CLI**：`npm install -g @anthropic-ai/claude-code && claude login`
- **WPS Office**（macOS 或 Windows）

### 1. 克隆并配置

```bash
git clone https://github.com/claude-office-skills/claude-wps-ppt-plugin.git
cd claude-wps-ppt-plugin
cp .env.example .env
```

编辑 `.env` 选择模型（默认 Sonnet 4.6）：

```
VITE_CLAUDE_MODEL=claude-sonnet-4-6
```

### 2. 安装依赖 & 启动

```bash
npm install
npm run start
```

### 3. 注册 WPS 加载项

```bash
chmod +x install-to-wps.sh
./install-to-wps.sh
```

重启 WPS Office，打开演示文稿 → 点击顶部 **Claude AI** 标签 → **打开 Claude**。

## 项目结构

```
├── AGENTS.md              # AI 行为准则（PPT 特化）
├── CHANGELOG.md           # 版本日志
├── proxy-server.js        # Express 代理服务器 (:3002)
├── wps-addon/
│   ├── main.js            # WPS Host（PPT 上下文采集 + 代码执行桥 + 逐页动画）
│   └── ribbon.xml         # Ribbon 按钮定义
├── src/
│   ├── App.tsx            # 主界面（PPT 适配版）
│   ├── api/
│   │   ├── claudeClient.ts   # Claude API + PPT 上下文
│   │   ├── wpsAdapter.ts     # PPT 上下文适配层
│   │   └── sessionStore.ts   # 会话持久化
│   └── components/        # UI 组件
├── skills/
│   ├── bundled/           # PPT 内置技能
│   │   ├── ppt-core-api/     # WPP API 参考
│   │   ├── slide-generation/ # 幻灯片生成（含 4 套配色 + 5 种布局）
│   │   └── layout-beautify/  # 布局美化（装饰元素 + 视觉层次）
│   ├── modes/             # Agent / Plan / Ask 模式
│   └── connectors/        # MCP 连接器
├── commands/              # 8 个快捷指令
└── install-to-wps.sh     # WPS 加载项注册脚本
```

## 自定义 Skills

在 `skills/bundled/` 下创建新目录，添加 `SKILL.md` 文件：

```yaml
---
name: my-custom-skill
description: 自定义能力描述
context:
  keywords: [关键词1, 关键词2]
---

## 规则和代码模式

在此编写 Skill 内容...
```

重启 Proxy 即可自动加载。

## 版本日志

### v1.1.0 (2026-02-27) — Ribbon 图标修复 + 模型选择器修复

**修复**
- Ribbon 图标从 SVG URL 改为 PNG 本地文件引用，提升兼容性
- 修复模型选择器下拉菜单无法展开的问题（`overflow: hidden` → `overflow: visible`）

### v1.0.0 (2026-02-26) — 初始版本

**核心功能**
- 逐页翻页动画：代码执行后自动逐页 GotoSlide，每页停留 ~2.4 秒
- 高品质 System Prompt：设计质量标准作为最高优先级
- slide-generation skill v2：4 套配色方案 + 5 种布局模板
- layout-beautify skill v2：装饰元素模式 + 5 级字号视觉层次
- ppt-core-api skill：WPP API 完整参考
- 8 个快捷指令 + 3 种交互模式 + MCP 连接器
- 代码执行桥 + 流式响应 + 会话历史
- 侧边栏默认左侧 + 崩溃恢复 + 橙色品牌加载动画

**修复**
- "已应用到表格" → "已应用到演示文稿"
- TASKPANE_URL 从 Vite dev server 修正为 proxy
- 侧边栏拖动崩溃后无法重新打开

## 相关项目

| 项目 | 说明 |
|------|------|
| [claude-wps-plugin](https://github.com/claude-office-skills/claude-wps-plugin) | Claude for WPS Excel |
| [claude-wps-word-plugin](https://github.com/claude-office-skills/claude-wps-word-plugin) | Claude for WPS Word |

## License

MIT
