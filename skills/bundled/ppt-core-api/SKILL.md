---
name: ppt-core-api
description: WPS WPP API 核心参考 — 全局变量、Slide/Shape 操作、文本框、图片、布局
version: 1.0.0
tags: [wps, api, ppt, slides, shapes, core]
modes: [agent, plan]
context:
  always: true
metadata:
  wps:
    minVersion: "6.0"
---

## 全局变量

- Application / app：WPS 应用对象
- Application.ActivePresentation：当前演示文稿
- Application.ActiveWindow：当前窗口
- Application.ActiveWindow.Selection：当前选区

## 核心 API（同步执行）

### 演示文稿

```javascript
var pres = Application.ActivePresentation;
var slideCount = pres.Slides.Count;
var slide = pres.Slides.Item(1);        // 1-based 索引
var pageW = pres.PageSetup.SlideWidth;  // 默认 960
var pageH = pres.PageSetup.SlideHeight; // 默认 540
```

### 添加幻灯片

```javascript
// 在末尾添加空白页
pres.Slides.Add(pres.Slides.Count + 1, 12);
// 参数2: Layout 枚举
//   1 = ppLayoutTitle (标题页)
//   2 = ppLayoutText (标题+正文)
//   6 = ppLayoutTitleOnly (仅标题)
//   7 = ppLayoutBlank (空白)
//   12 = ppLayoutCustom (自定义/空白)
var newSlide = pres.Slides.Item(pres.Slides.Count);
```

### 文本框

```javascript
// 添加文本框 (left, top, width, height)
var shape = slide.Shapes.AddTextbox(1, 50, 30, 860, 80);
shape.TextFrame.TextRange.Text = "标题文本";
shape.TextFrame.TextRange.Font.Size = 28;
shape.TextFrame.TextRange.Font.Bold = true;
shape.TextFrame.TextRange.Font.Color.RGB = 0x000000; // RGB 格式（非 BGR）
shape.TextFrame.TextRange.Font.Name = "微软雅黑";

// 段落对齐
shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2; // 1=左, 2=居中, 3=右

// 自动换行和内边距
shape.TextFrame.WordWrap = true;
shape.TextFrame.MarginLeft = 10;
shape.TextFrame.MarginTop = 5;
```

### 形状

```javascript
// 矩形 (left, top, width, height)
var rect = slide.Shapes.AddShape(1, 100, 200, 300, 150);
// 常用 AutoShapeType: 1=矩形, 5=圆角矩形, 9=圆形, 13=右箭头

rect.Fill.ForeColor.RGB = 0x4472C4;         // 填充色
rect.Fill.Transparency = 0;                  // 0=不透明, 1=全透明
rect.Line.ForeColor.RGB = 0x2F528F;         // 边框色
rect.Line.Weight = 1.5;                      // 边框宽度
rect.Line.Visible = true;

// 圆角矩形的圆角
rect.Adjustments.Item(1) = 0.2;              // 0~1，值越大越圆
```

### 图片

```javascript
// 插入图片
var pic = slide.Shapes.AddPicture(
  "/path/to/image.png",  // 文件路径
  false,                  // LinkToFile
  true,                   // SaveWithDocument
  100, 100,               // Left, Top
  400, 300                // Width, Height
);
```

### 位置和大小

```javascript
shape.Left = 100;     // 左边距（磅）
shape.Top = 200;      // 上边距（磅）
shape.Width = 400;    // 宽度（磅）
shape.Height = 300;   // 高度（磅）
shape.Rotation = 0;   // 旋转角度
```

### 删除

```javascript
shape.Delete();                    // 删除单个形状
slide.Delete();                    // 删除整页幻灯片
pres.Slides.Item(3).Delete();     // 删除第3页
```

### 选区操作

```javascript
var sel = Application.ActiveWindow.Selection;
// sel.Type: 0=None, 1=Slides, 2=Shapes, 3=Text
if (sel.ShapeRange) {
  for (var i = 1; i <= sel.ShapeRange.Count; i++) {
    var s = sel.ShapeRange.Item(i);
    // 操作选中形状
  }
}
```

### 备注

```javascript
slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = "演讲者备注内容";
```

## 不可用 API（严禁使用，会报错！）

- ❌ slide.Shapes.AddTable() 可能在某些 WPS 版本不可用，建议用文本框模拟表格
- ❌ shape.TextFrame2 — WPS 插件中不支持，必须用 shape.TextFrame
- ❌ Application.Presentations.Open() — 插件无法打开新文件
- ⚠️ 颜色使用 RGB 格式（0xRRGGBB），不是 BGR

## 常用颜色速查

| 用途 | RGB 值 | 说明 |
|------|--------|------|
| 深蓝(品牌色) | 0x4472C4 | 适合标题/强调 |
| 深灰(正文) | 0x333333 | 正文文字 |
| 浅灰(辅助) | 0x999999 | 副标题/说明 |
| 白色 | 0xFFFFFF | 背景/反色文字 |
| 橙色(强调) | 0xED7D31 | CTA/重点 |
| 绿色(正面) | 0x70AD47 | 增长/完成 |
| 红色(警告) | 0xFF0000 | 下降/警告 |

## 页面尺寸参考

- 标准 16:9: width=960, height=540
- 标准 4:3: width=720, height=540
- 宽屏: width=1280, height=720
