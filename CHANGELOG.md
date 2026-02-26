# Changelog

All notable changes to Claude for WPS PowerPoint will be documented in this file.

## [1.0.0] - 2026-02-26

### Added
- **逐页翻页动画**：代码执行后自动逐页 GotoSlide，每页停留 ~2.4 秒，呈现"AI 逐页创建"效果
- **高品质 System Prompt**：设计质量标准作为最高优先级，要求封面深色背景、卡片布局、装饰元素
- **slide-generation skill v2**：4 套配色方案（科技蓝/商务橙/自然绿/高端紫）+ 5 种布局模板（封面/目录/卡片/数据亮点/结尾）
- **layout-beautify skill v2**：装饰元素模式（色条/色带/侧边条）+ 5 级字号视觉层次体系
- **ppt-core-api skill**：WPP API 完整参考（演示文稿/幻灯片/形状/文本框/图片/颜色速查）
- **8 个快捷指令**：生成幻灯片、美化布局、添加备注、提取文本、翻译、编辑形状、添加图片、调整样式
- **3 种交互模式**：Agent（自动执行 maxTurns:3）/ Plan（步骤规划 maxTurns:8）/ Ask（只读分析）
- **MCP 连接器**：Web Search（Tavily）/ Knowledge Base
- **代码执行桥**：前端提交 → Proxy 队列 → Host 轮询执行 → 结果回传
- **流式响应**：SSE 流式输出 + Markdown 渲染 + 代码块语法高亮
- **会话历史**：自动保存对话记录，支持多会话切换
- **侧边栏默认左侧**：DockPosition 设为 Left
- **崩溃恢复**：TaskPane 崩溃后自动清除无效引用，重新创建面板
- **橙色加载动画**：品牌色 #FF6B35 与 PPT 主题统一

### Fixed
- "已应用到表格" → "已应用到演示文稿"（从 Excel 复制过来的错误文案）
- TASKPANE_URL 从 Vite dev server (5174) 修正为 proxy (3002)
- 侧边栏拖动崩溃后无法重新打开的问题
