/**
 * WPS PPT 加载项入口文件
 *
 * 1. Ribbon 按钮：打开 Claude 侧边栏 / JS 调试器
 * 2. 上下文同步：定时将 PPT 数据推送到 proxy-server
 * 3. 代码执行桥：轮询 proxy 的待执行代码队列并在 WPS PPT 上下文中执行
 */

var TASKPANE_URL = "http://127.0.0.1:3002/";
var PROXY_URL = "http://127.0.0.1:3002";
var CTX_INTERVAL = 2000;
var CODE_POLL_INTERVAL = 800;
var TASKPANE_KEY = "claude_ppt_taskpane_id";

var _ctxTimer = null;
var _codePollTimer = null;
var _slideAnimState = null;

// ── Ribbon 按钮回调 ──────────────────────────────────────────

function OnOpenClaudePanel() {
  try {
    var tsId = null;
    try {
      tsId = wps.PluginStorage.getItem(TASKPANE_KEY);
    } catch (e) {}

    if (tsId) {
      try {
        var existing = wps.GetTaskPane(tsId);
        if (existing) {
          try {
            existing.Visible = !existing.Visible;
            startBackgroundSync();
            return;
          } catch (visErr) {
            // TaskPane 引用已失效（如崩溃后），清除并重建
            try { wps.PluginStorage.setItem(TASKPANE_KEY, ""); } catch (e) {}
          }
        }
      } catch (e) {
        try { wps.PluginStorage.setItem(TASKPANE_KEY, ""); } catch (e2) {}
      }
    }

    var createFn = typeof wps.CreateTaskPane === "function"
      ? wps.CreateTaskPane
      : typeof wps.createTaskPane === "function"
        ? wps.createTaskPane
        : null;

    if (!createFn) {
      alert("当前 WPS 版本不支持 TaskPane API，请更新 WPS Office 到最新版本。");
      return;
    }

    var taskPane = createFn.call(wps, TASKPANE_URL);
    taskPane.DockPosition =
      wps.Enum && wps.Enum.JSKsoEnum_msoCTPDockPositionLeft
        ? wps.Enum.JSKsoEnum_msoCTPDockPositionLeft
        : 0;
    taskPane.Visible = true;

    try {
      wps.PluginStorage.setItem(TASKPANE_KEY, taskPane.ID);
    } catch (e) {}
    startBackgroundSync();
  } catch (e) {
    alert(
      "打开 Claude 面板失败：" +
        e.message +
        "\n\n请确保开发服务器已启动：\ncd ~/需求讨论/claude-wps-ppt-plugin && npm run dev",
    );
  }
}

function OnOpenJSDebugger() {
  try {
    if (
      typeof wps !== "undefined" &&
      wps.PluginStorage &&
      typeof wps.PluginStorage.openDebugger === "function"
    ) {
      wps.PluginStorage.openDebugger();
    } else if (typeof wps !== "undefined" && typeof wps.openDevTools === "function") {
      wps.openDevTools();
    } else if (
      typeof Application !== "undefined" &&
      Application.PluginStorage &&
      typeof Application.PluginStorage.openDebugger === "function"
    ) {
      Application.PluginStorage.openDebugger();
    } else {
      alert("JS 调试器在当前 WPS 版本下不可用。\n\n可尝试：菜单 → 开发工具 → 打开调试器");
    }
  } catch (e) {
    alert("打开调试器失败：" + e.message);
  }
}

function GetClaudeIcon() {
  return "claude-icon.png";
}

function GetDebugIcon() {
  return "debug-icon.png";
}

// ── 右键 "Add to Chat" ─────────────────────────────────────

function OnAddToChat() {
  try {
    var pres = Application.ActivePresentation;
    if (!pres) {
      alert("请先打开一个演示文稿");
      return;
    }

    var sel = Application.ActiveWindow.Selection;
    var texts = [];
    var shapeNames = [];
    var slideIndex = -1;

    try {
      if (sel.SlideRange && sel.SlideRange.Count > 0) {
        slideIndex = sel.SlideRange.Item(1).SlideIndex;
      }
    } catch (e) {}

    try {
      if (sel.ShapeRange && sel.ShapeRange.Count > 0) {
        for (var i = 1; i <= sel.ShapeRange.Count; i++) {
          var shape = sel.ShapeRange.Item(i);
          shapeNames.push(shape.Name || "Shape" + i);
          try {
            if (shape.TextFrame && shape.TextFrame.HasText) {
              texts.push(shape.TextFrame.TextRange.Text);
            }
          } catch (e) {}
        }
      }
    } catch (e) {}

    if (texts.length === 0 && slideIndex > 0) {
      try {
        var slide = pres.Slides.Item(slideIndex);
        for (var j = 1; j <= slide.Shapes.Count; j++) {
          try {
            var sh = slide.Shapes.Item(j);
            if (sh.TextFrame && sh.TextFrame.HasText) {
              texts.push(sh.TextFrame.TextRange.Text);
            }
          } catch (e) {}
        }
      } catch (e) {}
    }

    var payload = {
      type: "add-to-chat",
      slideIndex: slideIndex,
      shapeNames: shapeNames,
      texts: texts,
      timestamp: Date.now(),
    };

    httpPost(PROXY_URL + "/add-to-chat", JSON.stringify(payload));

    var tsId = null;
    try { tsId = wps.PluginStorage.getItem(TASKPANE_KEY); } catch (e) {}
    if (tsId) {
      try {
        var tp = wps.GetTaskPane(tsId);
        if (tp && !tp.Visible) tp.Visible = true;
      } catch (e) {}
    }
  } catch (e) {
    alert("Add to Chat 失败：" + e.message);
  }
}

function OnAddinLoad(ribbonUI) {
  if (typeof ribbonUI === "object") {
    // ribbon 引用
  }
  startBackgroundSync();
}

window.ribbon_bindUI = function (bindUI) {
  bindUI({
    OnOpenClaudePanel: OnOpenClaudePanel,
    OnAddToChat: OnAddToChat,
    OnOpenJSDebugger: OnOpenJSDebugger,
    GetClaudeIcon: GetClaudeIcon,
    GetDebugIcon: GetDebugIcon,
  });
};

// ── 后台同步启动 ─────────────────────────────────────────────

function startBackgroundSync() {
  if (_ctxTimer) {
    try { clearInterval(_ctxTimer); } catch (e) {}
  }
  if (_codePollTimer) {
    try { clearInterval(_codePollTimer); } catch (e) {}
  }
  pushWpsContext();
  _ctxTimer = setInterval(pushWpsContext, CTX_INTERVAL);
  _codePollTimer = setInterval(pollAndExecuteCode, CODE_POLL_INTERVAL);
}

// ── PPT 上下文推送 ───────────────────────────────────────────

function pushWpsContext() {
  try {
    var ctx = collectPptContext();
    if (!ctx.fileName) return;
    httpPost(PROXY_URL + "/wps-context", JSON.stringify(ctx));
  } catch (e) {}
}

function collectPptContext() {
  var result = {
    fileName: "",
    slideCount: 0,
    currentSlideIndex: null,
    pageSetup: { width: 0, height: 0 },
    selectedShapes: [],
    slideTexts: [],
  };

  try {
    var pres = Application.ActivePresentation;
    if (!pres) return result;

    result.fileName = pres.Name || "";
    result.slideCount = pres.Slides.Count;

    try {
      result.pageSetup = {
        width: pres.PageSetup.SlideWidth,
        height: pres.PageSetup.SlideHeight,
      };
    } catch (e) {}

    try {
      var sel = Application.ActiveWindow.Selection;
      if (sel.SlideRange && sel.SlideRange.Count > 0) {
        result.currentSlideIndex = sel.SlideRange.Item(1).SlideIndex;
      }
      result.selectedShapes = getSelectedShapesInfo(sel);
    } catch (e) {}

    result.slideTexts = getSlideTexts(pres.Slides);
  } catch (e) {}

  return result;
}

function getSelectedShapesInfo(sel) {
  var shapes = [];
  try {
    if (!sel.ShapeRange || sel.ShapeRange.Count === 0) return shapes;
    for (var i = 1; i <= sel.ShapeRange.Count; i++) {
      var shape = sel.ShapeRange.Item(i);
      var info = {
        id: shape.Id || i,
        name: shape.Name || "Shape" + i,
        type: getShapeTypeName(shape.Type),
        left: Math.round(shape.Left),
        top: Math.round(shape.Top),
        width: Math.round(shape.Width),
        height: Math.round(shape.Height),
      };
      try {
        if (shape.TextFrame && shape.TextFrame.HasText) {
          info.text = shape.TextFrame.TextRange.Text.substring(0, 500);
        }
      } catch (e) {}
      shapes.push(info);
    }
  } catch (e) {}
  return shapes;
}

function getShapeTypeName(typeVal) {
  var map = {
    1: "AutoShape",
    6: "Group",
    13: "Picture",
    14: "Placeholder",
    15: "MediaObject",
    17: "TextBox",
    19: "Table",
    21: "Chart",
    24: "SmartArt",
  };
  return map[typeVal] || "Shape(" + typeVal + ")";
}

function getSlideTexts(slides) {
  var result = [];
  var maxSlides = Math.min(slides.Count, 20);
  for (var i = 1; i <= maxSlides; i++) {
    var slide = slides.Item(i);
    var slideInfo = {
      index: i,
      layout: "",
      shapeCount: slide.Shapes.Count,
      texts: [],
    };
    try {
      slideInfo.layout = slide.Layout || "";
    } catch (e) {}
    for (var j = 1; j <= slide.Shapes.Count; j++) {
      try {
        var sh = slide.Shapes.Item(j);
        if (sh.TextFrame && sh.TextFrame.HasText) {
          var txt = sh.TextFrame.TextRange.Text;
          if (txt.length > 300) txt = txt.substring(0, 300) + "...";
          slideInfo.texts.push(txt);
        }
      } catch (e) {}
    }
    result.push(slideInfo);
  }
  return result;
}

// ── 代码执行桥 ───────────────────────────────────────────────

function pollAndExecuteCode() {
  if (_slideAnimState) {
    _processSlideAnim();
    return;
  }

  try {
    var resp = httpGet(PROXY_URL + "/pending-code");
    if (!resp) return;

    var data = JSON.parse(resp);
    if (!data.pending) return;

    var id = data.id;
    var code = data.code;

    var slideCountBefore = 0;
    try {
      slideCountBefore = Application.ActivePresentation.Slides.Count;
    } catch (e) {}

    try {
      var execResult = executeInWps(code);

      var slideCountAfter = 0;
      try {
        slideCountAfter = Application.ActivePresentation.Slides.Count;
      } catch (e) {}

      var newSlides = slideCountAfter - slideCountBefore;
      var totalSlides = slideCountAfter;
      var isBatchOp = code.length > 300 && totalSlides > 1;
      // #region agent log
      try{fetch('http://127.0.0.1:7244/ingest/63acb95d-6f91-4165-a07a-5bab2abb61eb',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'38bc8c'},body:JSON.stringify({sessionId:'38bc8c',location:'wps-addon/main.js:execDone',message:'PPT exec done',data:{before:slideCountBefore,after:slideCountAfter,newSlides:newSlides,codeLen:code.length,isBatchOp:isBatchOp,totalSlides:totalSlides},timestamp:Date.now(),hypothesisId:'H-anim2'})}).catch(function(){});}catch(e){}
      // #endregion
      if (newSlides > 1) {
        _startSlideAnim(slideCountBefore + 1, slideCountAfter, id, execResult);
      } else if (isBatchOp) {
        _startSlideAnim(1, totalSlides, id, execResult);
      } else {
        if (newSlides === 1) {
          try { Application.ActiveWindow.View.GotoSlide(slideCountAfter); } catch(e){}
        }
        httpPost(
          PROXY_URL + "/code-result",
          JSON.stringify({ id: id, result: execResult }),
        );
      }
    } catch (execErr) {
      httpPost(
        PROXY_URL + "/code-result",
        JSON.stringify({ id: id, error: execErr.message || String(execErr) }),
      );
    }
  } catch (e) {}
}

function _startSlideAnim(firstNew, lastSlide, id, result) {
  // #region agent log
  try{fetch('http://127.0.0.1:7244/ingest/63acb95d-6f91-4165-a07a-5bab2abb61eb',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'38bc8c'},body:JSON.stringify({sessionId:'38bc8c',location:'wps-addon/main.js:startAnim',message:'SLIDE ANIM START',data:{firstNew:firstNew,lastSlide:lastSlide},timestamp:Date.now(),hypothesisId:'H-anim3'})}).catch(function(){});}catch(e){}
  // #endregion
  try { Application.ActiveWindow.View.GotoSlide(firstNew); } catch(e){}
  _slideAnimState = {
    currentSlide: firstNew + 1,
    lastSlide: lastSlide,
    id: id,
    result: result,
    tickWait: 0
  };
}

function _processSlideAnim() {
  var st = _slideAnimState;
  if (!st) return;

  if (st.tickWait > 0) {
    st.tickWait--;
    return;
  }

  if (st.currentSlide > st.lastSlide) {
    httpPost(
      PROXY_URL + "/code-result",
      JSON.stringify({ id: st.id, result: st.result }),
    );
    _slideAnimState = null;
    return;
  }

  // #region agent log
  try{fetch('http://127.0.0.1:7244/ingest/63acb95d-6f91-4165-a07a-5bab2abb61eb',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'38bc8c'},body:JSON.stringify({sessionId:'38bc8c',location:'wps-addon/main.js:animStep',message:'SLIDE ANIM STEP',data:{goTo:st.currentSlide,lastSlide:st.lastSlide},timestamp:Date.now(),hypothesisId:'H-anim3'})}).catch(function(){});}catch(e){}
  // #endregion
  try {
    Application.ActiveWindow.View.GotoSlide(st.currentSlide);
  } catch (e) {}

  st.currentSlide++;
  st.tickWait = 2;
}

function executeInWps(code) {
  var fn = new Function(code);
  var result = fn();
  return result === undefined ? "执行成功" : String(result);
}

// ── HTTP 工具（同步 XHR）─────────────────────────────────────

function httpPost(url, body) {
  try {
    var xhr = new XMLHttpRequest();
    xhr.open("POST", url, false);
    xhr.setRequestHeader("Content-Type", "application/json");
    xhr.send(body);
    return xhr.responseText;
  } catch (e) {
    return null;
  }
}

function httpGet(url) {
  try {
    var xhr = new XMLHttpRequest();
    xhr.open("GET", url, false);
    xhr.send();
    return xhr.responseText;
  } catch (e) {
    return null;
  }
}
