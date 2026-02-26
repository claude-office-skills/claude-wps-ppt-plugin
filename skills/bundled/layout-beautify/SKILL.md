---
name: layout-beautify
description: 布局美化 — 专业排版、配色优化、视觉层次、装饰元素
version: 2.0.0
tags: [ppt, layout, beautify, alignment, color, font, design]
modes: [agent, plan]
context:
  keywords:
    - 美化
    - 对齐
    - 布局
    - 配色
    - 字体
    - 样式
    - 统一
    - 规范
    - 排版
    - 间距
    - 颜色
    - 格式化
    - 好看
    - 漂亮
    - 设计
    - 风格
---

## 布局美化规则

### 视觉层次（从大到小）

| 层级 | 字号 | 字重 | 颜色 | 用途 |
|------|------|------|------|------|
| 封面标题 | 40-48 | Bold | 白色/主色 | 封面页标题 |
| 页面标题 | 28-34 | Bold | 主色 | 每页顶部标题 |
| 副标题 | 20-24 | SemiBold | 深灰 | 分节标题 |
| 正文 | 14-18 | Regular | 中灰 0x444444 | 内容文字 |
| 辅助 | 11-13 | Regular | 浅灰 0x999999 | 页脚、图注 |

### 装饰元素模式

#### 页面标题区装饰（每个正文页推荐使用）
```javascript
// 标题下方的主色色条（视觉锚点）
var accent = slide.Shapes.AddShape(1, 50, 70, 60, 5);
accent.Fill.ForeColor.RGB = PRIMARY_COLOR;
accent.Line.Visible = false;
```

#### 页面顶部色带
```javascript
// 顶部细色带（贯穿整页宽度）
var topBar = slide.Shapes.AddShape(1, 0, 0, pageW, 8);
topBar.Fill.ForeColor.RGB = PRIMARY_COLOR;
topBar.Line.Visible = false;
```

#### 侧边装饰条
```javascript
// 左侧竖向色条
var sideBar = slide.Shapes.AddShape(1, 0, 0, 8, pageH);
sideBar.Fill.ForeColor.RGB = PRIMARY_COLOR;
sideBar.Line.Visible = false;
```

### 间距规范

- 页面边距：左右 60pt，上 40pt，下 30pt
- 标题与第一个内容元素间距：25-35pt
- 卡片之间间距：16-24pt
- 要点行间距：SpaceAfter = 8-12
- 段落间距：SpaceAfter = 16-20

### 对齐规则

- 同一行的元素必须顶部对齐（Top 相同）
- 多列布局的列宽相等，间距相等
- 标题统一左对齐（Left = 60）
- 数字/数据居中对齐（Alignment = 2）
- 页面标题 + 正文区保持相同左边距

### 字体规范

- **中文字体**：微软雅黑（首选）、思源黑体
- **英文/数字字体**：微软雅黑自带英文部分即可
- 全文字体统一，同一 PPT 中不超过 2 种字体
- 数字可以比旁边的文字大 2-4pt 增加视觉权重

### 形状样式

```javascript
// 标准卡片样式
function styleCard(shape, fillColor) {
  shape.Fill.ForeColor.RGB = fillColor;
  shape.Line.Visible = false;
  try { shape.Adjustments.Item(1) = 0.06; } catch(e) {}
}

// 标准圆形图标底
function styleCircle(shape, fillColor) {
  shape.Fill.ForeColor.RGB = fillColor;
  shape.Line.Visible = false;
}
```

### 禁止事项

- ❌ 同一页超过 3 种字体
- ❌ 同一页超过 4 种颜色
- ❌ 文字小于 11pt
- ❌ 内容紧贴页面边缘（最小边距 30pt）
- ❌ 纯文字没有任何装饰元素
- ❌ 封面页使用白色背景（封面必须有视觉冲击力）
