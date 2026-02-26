---
name: slide-generation
description: 幻灯片内容生成 — 大纲规划、视觉设计、多种布局模式、批量创建高品质演示文稿
version: 2.0.0
tags: [ppt, generation, slides, outline, batch, design]
modes: [agent, plan]
context:
  keywords:
    - 生成
    - 创建
    - 新建
    - 幻灯片
    - PPT
    - 演示
    - 大纲
    - 汇报
    - 报告
    - 课件
    - 培训
    - 制作
    - 模板
    - 年度
    - 季度
    - 工作
    - 总结
    - 计划
---

## 演示文稿设计规则

### 核心设计原则

1. **视觉层次**：每页必须有明确的视觉焦点，标题→副标题→正文→辅助信息
2. **留白呼吸**：内容不超过页面 60%，充分留白
3. **色彩统一**：每套 PPT 使用一个主题色系，最多 3 种主色
4. **形状装饰**：用圆角矩形卡片、色块、装饰线条提升视觉品质

### 配色方案（根据主题自动选择）

#### 科技蓝（适合：技术、产品、互联网）
| 元素 | RGB | 说明 |
|------|-----|------|
| 深蓝背景 | 0x0D1B2A | 封面/强调页背景 |
| 主蓝 | 0x1B4DFF | 标题、强调元素 |
| 亮蓝 | 0x4CC9F0 | 图标、装饰 |
| 浅蓝底 | 0xE8F4FD | 卡片背景 |
| 白色 | 0xFFFFFF | 文字（深色背景上）|

#### 商务橙（适合：年度汇报、营销、商业计划）
| 元素 | RGB | 说明 |
|------|-----|------|
| 深灰背景 | 0x1A1A2E | 封面背景 |
| 主橙 | 0xFF6B35 | 标题、数据亮点 |
| 暖橙 | 0xFFA062 | 辅助元素 |
| 浅橙底 | 0xFFF3ED | 卡片背景 |
| 深灰 | 0x2D2D2D | 正文 |

#### 自然绿（适合：教育、环保、健康、农业）
| 元素 | RGB | 说明 |
|------|-----|------|
| 深绿背景 | 0x0B3D2E | 封面背景 |
| 主绿 | 0x27AE60 | 标题、图标 |
| 亮绿 | 0x6BCB77 | 辅助装饰 |
| 浅绿底 | 0xEDF7F0 | 卡片背景 |
| 深灰 | 0x333333 | 正文 |

#### 高端紫（适合：文化、创意、品牌、奢侈品）
| 元素 | RGB | 说明 |
|------|-----|------|
| 深紫背景 | 0x1A0A2E | 封面背景 |
| 主紫 | 0x7C3AED | 标题、强调 |
| 亮紫 | 0xA78BFA | 辅助元素 |
| 浅紫底 | 0xF3EEFF | 卡片背景 |
| 深灰 | 0x2D2D2D | 正文 |

### 页面结构模板

#### 封面页（必须有色彩冲击力）
```javascript
// 深色全屏背景 + 大标题 + 副标题
var slide = pres.Slides.Item(idx);
// 背景色块
var bg = slide.Shapes.AddShape(1, 0, 0, pageW, pageH);
bg.Fill.ForeColor.RGB = BG_COLOR;
bg.Line.Visible = false;
// 装饰线
var line = slide.Shapes.AddShape(1, pageW * 0.1, pageH * 0.55, pageW * 0.15, 4);
line.Fill.ForeColor.RGB = ACCENT_COLOR;
line.Line.Visible = false;
// 主标题（大号白字）
var t = slide.Shapes.AddTextbox(1, pageW * 0.1, pageH * 0.25, pageW * 0.8, 80);
t.TextFrame.TextRange.Text = title;
t.TextFrame.TextRange.Font.Size = 44;
t.TextFrame.TextRange.Font.Bold = true;
t.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF;
t.TextFrame.TextRange.Font.Name = "微软雅黑";
// 副标题
var sub = slide.Shapes.AddTextbox(1, pageW * 0.1, pageH * 0.6, pageW * 0.8, 40);
sub.TextFrame.TextRange.Text = subtitle;
sub.TextFrame.TextRange.Font.Size = 18;
sub.TextFrame.TextRange.Font.Color.RGB = 0xCCCCCC;
```

#### 目录页（圆形编号 + 标题列表）
```javascript
// 每个章节：圆形编号 + 标题文字
for (var i = 0; i < chapters.length; i++) {
  var y = 120 + i * 90;
  // 编号圆形
  var circle = slide.Shapes.AddShape(9, 80, y, 50, 50);
  circle.Fill.ForeColor.RGB = PRIMARY_COLOR;
  circle.Line.Visible = false;
  var numTxt = circle.TextFrame.TextRange;
  numTxt.Text = "0" + (i + 1);
  numTxt.Font.Size = 18;
  numTxt.Font.Color.RGB = 0xFFFFFF;
  numTxt.Font.Bold = true;
  numTxt.ParagraphFormat.Alignment = 2;
  // 标题文字
  var chTitle = slide.Shapes.AddTextbox(1, 150, y + 8, 600, 36);
  chTitle.TextFrame.TextRange.Text = chapters[i];
  chTitle.TextFrame.TextRange.Font.Size = 22;
  chTitle.TextFrame.TextRange.Font.Color.RGB = 0x333333;
}
```

#### 内容页 — 卡片布局（2~4 列）
```javascript
// 卡片式布局：圆角矩形 + 标题 + 描述
var cardW = (pageW - 120) / cols - 20;
for (var i = 0; i < items.length; i++) {
  var col = i % cols;
  var x = 60 + col * (cardW + 20);
  var y = 130;
  // 卡片背景（圆角矩形）
  var card = slide.Shapes.AddShape(5, x, y, cardW, 340);
  card.Fill.ForeColor.RGB = CARD_BG;
  card.Line.Visible = false;
  try { card.Adjustments.Item(1) = 0.08; } catch(e){}
  // 卡片内标题
  var ct = slide.Shapes.AddTextbox(1, x + 20, y + 20, cardW - 40, 40);
  ct.TextFrame.TextRange.Text = items[i].title;
  ct.TextFrame.TextRange.Font.Size = 20;
  ct.TextFrame.TextRange.Font.Bold = true;
  ct.TextFrame.TextRange.Font.Color.RGB = PRIMARY_COLOR;
  // 卡片内描述
  var cd = slide.Shapes.AddTextbox(1, x + 20, y + 70, cardW - 40, 250);
  cd.TextFrame.TextRange.Text = items[i].desc;
  cd.TextFrame.TextRange.Font.Size = 14;
  cd.TextFrame.TextRange.Font.Color.RGB = 0x666666;
  cd.TextFrame.WordWrap = true;
}
```

#### 内容页 — 数据亮点（大数字 + 说明）
```javascript
// 大数字高亮展示
for (var i = 0; i < metrics.length; i++) {
  var x = 60 + i * (pageW / metrics.length);
  var w = pageW / metrics.length - 40;
  // 数字（超大号 + 主色）
  var num = slide.Shapes.AddTextbox(1, x, 180, w, 80);
  num.TextFrame.TextRange.Text = metrics[i].value;
  num.TextFrame.TextRange.Font.Size = 56;
  num.TextFrame.TextRange.Font.Bold = true;
  num.TextFrame.TextRange.Font.Color.RGB = PRIMARY_COLOR;
  num.TextFrame.TextRange.ParagraphFormat.Alignment = 2;
  // 说明（小字灰色）
  var label = slide.Shapes.AddTextbox(1, x, 270, w, 30);
  label.TextFrame.TextRange.Text = metrics[i].label;
  label.TextFrame.TextRange.Font.Size = 16;
  label.TextFrame.TextRange.Font.Color.RGB = 0x999999;
  label.TextFrame.TextRange.ParagraphFormat.Alignment = 2;
}
```

#### 结尾页（简洁有力）
```javascript
// 深色背景 + 致谢 + 联系方式
var bg = slide.Shapes.AddShape(1, 0, 0, pageW, pageH);
bg.Fill.ForeColor.RGB = BG_COLOR;
bg.Line.Visible = false;
var thanks = slide.Shapes.AddTextbox(1, 0, pageH * 0.35, pageW, 60);
thanks.TextFrame.TextRange.Text = "感谢聆听";
thanks.TextFrame.TextRange.Font.Size = 40;
thanks.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF;
thanks.TextFrame.TextRange.Font.Bold = true;
thanks.TextFrame.TextRange.ParagraphFormat.Alignment = 2;
```

### 内容生成规则

1. **标题简洁有力**：每页标题不超过 10 个字
2. **要点精炼**：每个要点不超过 20 字，每页不超过 5 个要点
3. **使用数据说话**：能用数字就不用文字（如"增长 35%"而不是"显著增长"）
4. **层次分明**：大标题→小标题→要点→补充说明
5. **专业术语恰当**：根据用户行业使用对应专业词汇

### 代码规范

- 生成超过 3 页时**必须**使用数据数组 + 循环
- 所有颜色定义为变量（顶部声明），便于主题统一
- 使用 `pageW` 和 `pageH` 做相对布局，不硬编码绝对位置
- 每页必须有标题区和装饰元素（色块/线条/形状）
- 封面和结尾页**必须**使用深色背景
- 正文页使用白色或浅色背景 + 卡片布局
