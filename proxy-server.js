/**
 * 本地代理服务器（PPT 版）
 *
 * 1) 接收浏览器插件的请求，调用本地已认证的 claude CLI 执行，以 SSE 流式返回响应。
 * 2) WPS PPT 上下文中转：Plugin Host POST 数据，Task Pane GET 读取。
 * 3) 代码执行桥：Task Pane 提交代码 → proxy 存入队列 → Plugin Host 轮询执行 → 结果回传。
 *
 * 运行：node proxy-server.js
 * 端口：3002
 */
import express from "express";
import cors from "cors";
import { spawn, execSync } from "child_process";
import {
  readFileSync,
  writeFileSync,
  mkdirSync,
  readdirSync,
  existsSync,
  unlinkSync,
} from "fs";
import { tmpdir } from "os";
import { join, dirname, resolve } from "path";
import { fileURLToPath } from "url";
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const pdfParse = require("pdf-parse");
const yaml = require("js-yaml");

const __dirname = dirname(fileURLToPath(import.meta.url));

function isPathSafe(filePath, allowedDir) {
  const resolved = resolve(filePath);
  const allowed = resolve(allowedDir);
  return resolved.startsWith(allowed + "/") || resolved === allowed;
}

const SESSION_ID_RE = /^[a-zA-Z0-9_-]+$/;

// ── Skill Loader ──────────────────────────────────────────────
function loadSkillsFromDir(subDir) {
  const skillsDir = join(__dirname, "skills", subDir);
  const skills = new Map();

  if (!existsSync(skillsDir)) return skills;

  for (const dir of readdirSync(skillsDir, { withFileTypes: true })) {
    if (!dir.isDirectory()) continue;
    const skillFile = join(skillsDir, dir.name, "SKILL.md");
    if (!existsSync(skillFile)) continue;

    const raw = readFileSync(skillFile, "utf-8");
    const { frontmatter, body } = parseFrontmatter(raw);
    skills.set(dir.name, {
      ...frontmatter,
      body,
      name: frontmatter.name || dir.name,
    });
  }

  return skills;
}

function loadSkills() {
  return loadSkillsFromDir("bundled");
}

function parseFrontmatter(raw) {
  const match = raw.match(/^---\n([\s\S]*?)\n---\n([\s\S]*)$/);
  if (!match) return { frontmatter: {}, body: raw };

  try {
    const fm = yaml.load(match[1]) || {};
    return { frontmatter: fm, body: match[2].trim() };
  } catch {
    return { frontmatter: {}, body: match[2].trim() };
  }
}

function matchSkills(allSkills, userMessage, pptContext, mode) {
  const matched = [];
  for (const [id, skill] of allSkills) {
    if (mode && Array.isArray(skill.modes) && !skill.modes.includes(mode)) {
      continue;
    }

    const ctx = skill.context || {};
    if (ctx.always === true || ctx.always === "true") {
      matched.push(skill);
      continue;
    }

    let keywordHit = false;
    if (Array.isArray(ctx.keywords)) {
      const msg = userMessage.toLowerCase();
      if (ctx.keywords.some((kw) => msg.includes(kw.toLowerCase()))) {
        keywordHit = true;
      }
    }

    let contextHit = false;
    if (pptContext && pptContext.selectedShapes) {
      if (
        (ctx.hasSelectedShapes === true || ctx.hasSelectedShapes === "true") &&
        pptContext.selectedShapes.length > 0
      )
        contextHit = true;
      if (
        ctx.minSlides &&
        pptContext.slideCount >= Number(ctx.minSlides)
      )
        contextHit = true;
    }

    if (keywordHit || contextHit) {
      matched.push(skill);
    }
  }
  return matched;
}

function buildSystemPrompt(skills, todayStr, userMessage, modeSkill) {
  let prompt = `你是 Claude，嵌入在 WPS Office PowerPoint 中的 AI 演示文稿助手。你生成的代码直接运行在 WPS Plugin Host 上下文，可同步访问完整 WPP API。\n今天的日期是 ${todayStr}。

## ⚠️ 设计质量标准（最高优先级）
你生成的演示文稿必须达到**专业设计师水准**：
- **封面页和结尾页**必须使用深色背景（如 0x0D1B2A, 0x1A1A2E），白色大标题，创造视觉冲击力
- **正文页**使用白色/浅色背景 + 卡片式布局（圆角矩形色块 + 图标/编号）
- **每页必须有装饰元素**：标题下方色条、顶部色带、编号圆形等
- **配色统一**：顶部声明颜色变量，全文统一使用，根据主题选择色系
- **排版精致**：使用 pageW/pageH 做相对布局，保持对齐和留白
- **内容精炼**：标题 ≤10 字，每个要点 ≤20 字，能用数字就不用文字
- **禁止纯白底纯文字**：每页都要有形状元素提升视觉品质

## ⚠️ 上下文优先级
每次请求都会附带「当前 PPT 上下文」，其中包含演示文稿名称、当前页码、选中形状和各页文本。
- 优先操作当前页，除非用户明确指定其他页面
- 如果用户切换了页面，以最新上下文中的页码为准
- 生成的代码必须基于 Application.ActivePresentation 操作

## ⚠️ 代码规范（必须遵守）
- 代码总长度 ≤ 5000 字符
- 批量生成时必须使用**数据数组 + 循环**
- 颜色值顶部声明为变量（如 var PRIMARY = 0x1B4DFF;）
- 使用 var pageW = pres.PageSetup.SlideWidth; 做相对布局
- 必须用 var pres = Application.ActivePresentation; 不要硬编码文件名

\n`;

  if (modeSkill && modeSkill.body) {
    prompt += modeSkill.body + "\n\n";
  }

  for (const skill of skills) {
    prompt += skill.body + "\n\n";
  }

  return prompt;
}

const ALL_SKILLS = loadSkills();
const ALL_MODES = loadSkillsFromDir("modes");
const ALL_CONNECTORS = loadSkillsFromDir("connectors");
const ALL_WORKFLOWS = loadSkillsFromDir("workflows");

console.log(
  `[skill-loader] bundled: ${ALL_SKILLS.size} (${[...ALL_SKILLS.keys()].join(", ")})`,
);
console.log(
  `[skill-loader] modes: ${ALL_MODES.size} (${[...ALL_MODES.keys()].join(", ")})`,
);
console.log(
  `[skill-loader] connectors: ${ALL_CONNECTORS.size}, workflows: ${ALL_WORKFLOWS.size}`,
);

// ── Command Loader ────────────────────────────────────────────
function loadCommands() {
  const cmdsDir = join(__dirname, "commands");
  const commands = [];

  if (!existsSync(cmdsDir)) return commands;

  for (const file of readdirSync(cmdsDir)) {
    if (!file.endsWith(".md")) continue;
    const raw = readFileSync(join(cmdsDir, file), "utf-8");
    const { frontmatter, body } = parseFrontmatter(raw);
    commands.push({
      id: file.replace(/\.md$/, ""),
      icon: frontmatter.icon || "📌",
      label: frontmatter.label || file.replace(/\.md$/, ""),
      description: frontmatter.description || "",
      scope: frontmatter.scope || "general",
      prompt: body.trim(),
    });
  }

  return commands;
}

const ALL_COMMANDS = loadCommands();
console.log(`[command-loader] 已加载 ${ALL_COMMANDS.length} 个 commands`);

const app = express();
const PORT = 3002;

app.use(
  cors({
    origin: [
      "http://127.0.0.1:3002",
      "http://localhost:3002",
      "http://127.0.0.1:5174",
      "http://localhost:5174",
    ],
  }),
);
app.use(express.json({ limit: "50mb" }));

const distPath = join(__dirname, "dist");
if (existsSync(distPath)) {
  app.use(express.static(distPath));
}

const wpsAddonPath = join(__dirname, "wps-addon");
if (existsSync(wpsAddonPath)) {
  app.use("/wps-addon", express.static(wpsAddonPath));
}

// ── 系统剪贴板读取（macOS）─────────────────────────────────────
app.get("/clipboard", (req, res) => {
  try {
    let hasImage = false;
    try {
      const types = execSync(
        `osascript -e 'clipboard info' 2>/dev/null | head -5`,
        { encoding: "utf-8", timeout: 2000 },
      );
      hasImage = /TIFF|PNG|JPEG|picture/i.test(types);
    } catch {}

    if (hasImage) {
      try {
        const imgName = `clipboard-${Date.now()}.png`;
        const imgPath = join(TEMP_DIR, imgName);
        execSync(
          `osascript -e 'set pngData to (the clipboard as «class PNGf»)' -e 'set fp to open for access POSIX file "${imgPath}" with write permission' -e 'write pngData to fp' -e 'close access fp'`,
          { timeout: 5000 },
        );
        return res.json({
          ok: true,
          type: "image",
          filePath: imgPath,
          fileName: imgName,
        });
      } catch {}
    }

    const text = execSync("pbpaste", {
      encoding: "utf-8",
      timeout: 2000,
      maxBuffer: 1024 * 1024,
      env: { ...process.env, LANG: "en_US.UTF-8" },
    });
    res.json({ ok: true, type: "text", text });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ── PDF 文本提取 ──────────────────────────────────────────────
app.post("/extract-pdf", async (req, res) => {
  try {
    const { base64, filePath } = req.body;
    let buffer;

    if (filePath) {
      if (!isPathSafe(filePath, TEMP_DIR)) {
        return res.status(400).json({ ok: false, error: "filePath 不合法" });
      }
      buffer = readFileSync(filePath);
    } else if (base64) {
      buffer = Buffer.from(base64, "base64");
    } else {
      return res
        .status(400)
        .json({ ok: false, error: "需要 base64 或 filePath" });
    }

    const uint8 = new Uint8Array(buffer);
    const parser = new pdfParse.PDFParse(uint8);
    const data = await parser.getText();
    const text = data.text || "";
    const pages = data.total || data.pages?.length || 0;

    const MAX_CHARS = 100000;
    const truncated = text.length > MAX_CHARS;
    const content = truncated ? text.slice(0, MAX_CHARS) : text;

    res.json({
      ok: true,
      text: content,
      pages,
      totalChars: text.length,
      truncated,
    });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ── 图片临时文件上传 ─────────────────────────────────────────
const TEMP_DIR = join(tmpdir(), "wps-claude-ppt-uploads");
try {
  mkdirSync(TEMP_DIR, { recursive: true });
} catch {}

let _tempFileCounter = 0;

app.post("/upload-temp", (req, res) => {
  try {
    const { base64, fileName } = req.body;
    if (!base64 || !fileName) {
      return res
        .status(400)
        .json({ ok: false, error: "需要 base64 和 fileName" });
    }
    const ext = fileName.includes(".")
      ? fileName.slice(fileName.lastIndexOf("."))
      : ".bin";
    const safeName = `upload-${++_tempFileCounter}-${Date.now()}${ext}`;
    const filePath = join(TEMP_DIR, safeName);
    writeFileSync(filePath, Buffer.from(base64, "base64"));
    res.json({ ok: true, filePath, fileName: safeName });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ── Commands API ──────────────────────────────────────────────
app.get("/commands", (req, res) => {
  const scope = req.query.scope;
  const filtered = scope
    ? ALL_COMMANDS.filter((c) => c.scope === scope)
    : ALL_COMMANDS;
  res.json(filtered);
});

// ── Skills 列表 API ─────────────────────────────────────────
app.get("/skills", (req, res) => {
  const list = [...ALL_SKILLS.entries()].map(([id, s]) => ({
    id,
    name: s.name,
    description: s.description,
    tags: s.tags,
    context: s.context,
  }));
  res.json(list);
});

app.get("/health", (req, res) => {
  // #region agent log
  fetch('http://127.0.0.1:7244/ingest/63acb95d-6f91-4165-a07a-5bab2abb61eb',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'38bc8c'},body:JSON.stringify({sessionId:'38bc8c',location:'proxy-server.js:health',message:'PPT health check hit',data:{port:PORT,version:'1.0.0-ppt'},timestamp:Date.now(),hypothesisId:'H1'})}).catch(()=>{});
  // #endregion
  res.json({
    status: "ok",
    version: "1.0.0-ppt",
    skills: ALL_SKILLS.size,
    modes: ALL_MODES.size,
    connectors: ALL_CONNECTORS.size,
    workflows: ALL_WORKFLOWS.size,
    commands: ALL_COMMANDS.length,
    skillNames: [...ALL_SKILLS.keys()],
    modeNames: [...ALL_MODES.keys()],
  });
});

app.get("/modes", (_req, res) => {
  const modes = [];
  for (const [id, skill] of ALL_MODES) {
    modes.push({
      id,
      name: skill.name,
      description: skill.description,
      default: skill.default === true || skill.default === "true",
      enforcement: skill.enforcement || {},
      quickActions: skill.quickActions || [],
    });
  }
  res.json(modes);
});

// ── 会话历史存储 ─────────────────────────────────────────────
const HISTORY_DIR = join(__dirname, ".chat-history");
const MEMORY_FILE = join(HISTORY_DIR, "memory.json");
try {
  mkdirSync(HISTORY_DIR, { recursive: true });
} catch {}

function loadMemory() {
  try {
    return JSON.parse(readFileSync(MEMORY_FILE, "utf-8"));
  } catch {
    return {
      preferences: {},
      frequentActions: [],
      lastModel: "claude-sonnet-4-6",
    };
  }
}

function saveMemory(mem) {
  writeFileSync(MEMORY_FILE, JSON.stringify(mem, null, 2));
}

app.get("/sessions", (req, res) => {
  try {
    if (!existsSync(HISTORY_DIR)) return res.json([]);
    const files = readdirSync(HISTORY_DIR)
      .filter((f) => f.endsWith(".json") && f !== "memory.json")
      .map((f) => {
        try {
          const data = JSON.parse(readFileSync(join(HISTORY_DIR, f), "utf-8"));
          return {
            id: data.id,
            title: data.title || "未命名会话",
            createdAt: data.createdAt,
            updatedAt: data.updatedAt,
            messageCount: data.messages?.length || 0,
            preview:
              data.messages
                ?.find((m) => m.role === "user")
                ?.content?.slice(0, 60) || "",
          };
        } catch {
          return null;
        }
      })
      .filter(Boolean)
      .sort((a, b) => b.updatedAt - a.updatedAt);
    res.json(files);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/sessions/:id", (req, res) => {
  try {
    if (!SESSION_ID_RE.test(req.params.id)) {
      return res.status(400).json({ error: "无效的会话 ID" });
    }
    const filePath = join(HISTORY_DIR, `${req.params.id}.json`);
    if (!existsSync(filePath))
      return res.status(404).json({ error: "会话不存在" });
    const data = JSON.parse(readFileSync(filePath, "utf-8"));
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post("/sessions", (req, res) => {
  try {
    const { id, title, messages, model } = req.body;
    if (!id) return res.status(400).json({ error: "id 不能为空" });
    if (!SESSION_ID_RE.test(id)) {
      return res.status(400).json({ error: "无效的会话 ID" });
    }
    const now = Date.now();
    const filePath = join(HISTORY_DIR, `${id}.json`);

    let session;
    if (existsSync(filePath)) {
      session = JSON.parse(readFileSync(filePath, "utf-8"));
      session.messages = messages || session.messages;
      session.title = title || session.title;
      session.model = model || session.model;
      session.updatedAt = now;
    } else {
      session = {
        id,
        title: title || "新会话",
        messages: messages || [],
        model,
        createdAt: now,
        updatedAt: now,
      };
    }

    writeFileSync(filePath, JSON.stringify(session, null, 2));

    const mem = loadMemory();
    if (model) mem.lastModel = model;
    saveMemory(mem);

    res.json({
      ok: true,
      session: {
        id: session.id,
        title: session.title,
        updatedAt: session.updatedAt,
      },
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete("/sessions/:id", (req, res) => {
  try {
    if (!SESSION_ID_RE.test(req.params.id)) {
      return res.status(400).json({ error: "无效的会话 ID" });
    }
    const filePath = join(HISTORY_DIR, `${req.params.id}.json`);
    if (existsSync(filePath)) {
      unlinkSync(filePath);
    }
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/memory", (req, res) => {
  res.json(loadMemory());
});

app.post("/memory", (req, res) => {
  try {
    const mem = loadMemory();
    Object.assign(mem, req.body);
    saveMemory(mem);
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── WPS PPT 上下文中转 ─────────────────────────────────────
let _pptContext = {
  fileName: "",
  slideCount: 0,
  currentSlideIndex: null,
  pageSetup: { width: 960, height: 540 },
  selectedShapes: [],
  slideTexts: [],
  timestamp: 0,
};

app.post("/wps-context", (req, res) => {
  if (!req.body.fileName && _pptContext.fileName) {
    res.json({ ok: true, skipped: true });
    return;
  }
  _pptContext = { ...req.body, timestamp: Date.now() };
  res.json({ ok: true });
});

app.get("/wps-context", (req, res) => {
  res.json(_pptContext);
});

// ── 代码执行桥 ─────────────────────────────────────────────
let _codeQueue = [];
let _codeResults = {};
let _codeIdCounter = 0;

app.post("/execute-code", (req, res) => {
  const { code } = req.body;
  if (!code) return res.status(400).json({ error: "code 不能为空" });

  const id = `exec-${++_codeIdCounter}-${Date.now()}`;
  _codeQueue.push({ id, code, submittedAt: Date.now() });
  res.json({ ok: true, id });
});

app.get("/pending-code", (req, res) => {
  if (_codeQueue.length === 0) {
    return res.json({ pending: false });
  }
  const item = _codeQueue.shift();
  res.json({ pending: true, ...item });
});

app.post("/code-result", (req, res) => {
  const { id, result, error } = req.body;
  if (!id) return res.status(400).json({ error: "id 不能为空" });

  _codeResults[id] = {
    result: result ?? null,
    error: error ?? null,
    completedAt: Date.now(),
  };

  setTimeout(() => {
    delete _codeResults[id];
  }, 60000);
  res.json({ ok: true });
});

app.get("/code-result/:id", (req, res) => {
  const entry = _codeResults[req.params.id];
  if (!entry) return res.json({ ready: false });
  res.json({ ready: true, ...entry });
});

// ── Add to Chat（右键菜单数据） ────────────────────────────
let _addToChatQueue = [];

app.post("/add-to-chat", (req, res) => {
  _addToChatQueue.push({ ...req.body, receivedAt: Date.now() });
  if (_addToChatQueue.length > 10) _addToChatQueue.shift();
  res.json({ ok: true });
});

app.get("/add-to-chat/poll", (_req, res) => {
  if (_addToChatQueue.length === 0) {
    return res.json({ pending: false });
  }
  const item = _addToChatQueue.shift();
  res.json({ pending: true, ...item });
});

// ── 模型白名单 ──────────────────────────────────────────────
const ALLOWED_MODELS = new Set([
  "claude-sonnet-4-6",
  "claude-opus-4-6",
  "claude-haiku-4-5",
]);

// ── 聊天接口（SSE 流式响应）──────────────────────────────────
app.post("/chat", (req, res) => {
  const { messages, context, model, attachments, webSearch, mode } = req.body;

  if (!messages || !Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: "messages 不能为空" });
  }

  const selectedModel = ALLOWED_MODELS.has(model) ? model : "claude-sonnet-4-6";

  const currentMode = mode || "agent";
  const modeSkill = ALL_MODES.get(currentMode) || ALL_MODES.get("agent");
  const enforcement = modeSkill?.enforcement || {};
  const skipCodeBridge =
    enforcement.codeBridge === false || enforcement.codeBridge === "false";

  const lastUserMsg = messages[messages.length - 1]?.content || "";
  const todayStr = new Date().toISOString().split("T")[0];
  const matchedSkills = matchSkills(ALL_SKILLS, lastUserMsg, _pptContext, currentMode);

  const matchedConnectors = matchSkills(
    ALL_CONNECTORS,
    lastUserMsg,
    _pptContext,
    currentMode,
  );
  const allMatched = [...matchedSkills, ...matchedConnectors];

  let fullPrompt =
    buildSystemPrompt(allMatched, todayStr, lastUserMsg, modeSkill) + "\n";

  const memory = loadMemory();
  if (memory.preferences && Object.keys(memory.preferences).length > 0) {
    fullPrompt += `[用户偏好记忆]\n`;
    for (const [k, v] of Object.entries(memory.preferences)) {
      fullPrompt += `- ${k}: ${v}\n`;
    }
    fullPrompt += "\n";
  }

  if (context) {
    fullPrompt += `[当前 PPT 上下文]\n${context}\n\n`;
  }

  if (Array.isArray(attachments) && attachments.length > 0) {
    const textAtts = attachments.filter((a) => a.type !== "image");
    const imageAtts = attachments.filter((a) => a.type === "image");

    if (textAtts.length > 0) {
      fullPrompt += "[用户附件]\n";
      textAtts.forEach((att) => {
        fullPrompt += `--- ${att.name} ---\n${att.content}\n\n`;
      });
    }

    if (imageAtts.length > 0) {
      fullPrompt += `[用户上传了 ${imageAtts.length} 张图片]\n`;
      imageAtts.forEach((att) => {
        if (att.tempPath) {
          if (!isPathSafe(att.tempPath, TEMP_DIR)) {
            fullPrompt += `图片 ${att.name}: 路径无效，已跳过\n`;
            return;
          }
          try {
            const imgBuf = readFileSync(att.tempPath);
            const ext = att.name?.split(".").pop()?.toLowerCase() || "png";
            const mime =
              {
                jpg: "jpeg",
                jpeg: "jpeg",
                png: "png",
                gif: "gif",
                webp: "webp",
                bmp: "bmp",
                svg: "svg+xml",
              }[ext] || "png";
            const b64 = imgBuf.toString("base64");
            fullPrompt += `图片 ${att.name}: data:image/${mime};base64,${b64.substring(0, 200)}... (${imgBuf.length} bytes, 已作为附件传入)\n`;
          } catch (e) {
            fullPrompt += `图片 ${att.name}: 无法读取 (${e.message})\n`;
          }
        }
      });
      fullPrompt +=
        "请根据图片和用户指令完成任务。如果用户要求参考图片创建幻灯片，请尽量还原图片中的布局和内容。\n\n";
    }
  }

  if (messages.length > 1) {
    fullPrompt += "[对话历史]\n";
    messages.slice(0, -1).forEach((m) => {
      const role = m.role === "user" ? "用户" : "助手";
      fullPrompt += `${role}: ${m.content}\n\n`;
    });
  }

  const lastMsg = messages[messages.length - 1];
  fullPrompt += `用户: ${lastMsg.content}`;

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");
  res.flushHeaders();

  const keepalive = setInterval(() => {
    if (!res.writableEnded) {
      res.write(`: keepalive ${Date.now()}\n\n`);
    }
  }, 5000);

  const claudePath = process.env.CLAUDE_PATH || "claude";
  const maxTurns = String(
    enforcement.maxTurns || (currentMode === "ask" ? 1 : 5),
  );
  const cliArgs = [
    "-p",
    "--verbose",
    "--output-format",
    "stream-json",
    "--include-partial-messages",
    "--max-turns",
    maxTurns,
    "--model",
    selectedModel,
  ];

  if (webSearch) {
    cliArgs.push("--allowedTools", "WebSearch");
  }
  const child = spawn(claudePath, cliArgs, { env: { ...process.env } });

  res.write(
    `data: ${JSON.stringify({ type: "mode", mode: currentMode, enforcement })}\n\n`,
  );

  child.stdin.write(fullPrompt);
  child.stdin.end();

  let resultText = "";
  let responseDone = false;
  let _lineBuf = "";
  let _tokenCount = 0;
  let _thinkingText = "";

  child.stdout.on("data", (data) => {
    _lineBuf += data.toString();
    const lines = _lineBuf.split("\n");
    _lineBuf = lines.pop() ?? "";

    for (const line of lines) {
      if (!line.trim()) continue;

      try {
        const evt = JSON.parse(line);

        if (evt.type === "stream_event") {
          const se = evt.event;

          if (se.type === "content_block_delta") {
            if (se.delta?.type === "text_delta" && se.delta.text) {
              resultText += se.delta.text;
              _tokenCount++;
              res.write(
                `data: ${JSON.stringify({ type: "token", text: se.delta.text })}\n\n`,
              );
            } else if (
              se.delta?.type === "thinking_delta" &&
              se.delta.thinking
            ) {
              _thinkingText += se.delta.thinking;
              res.write(
                `data: ${JSON.stringify({ type: "thinking", text: se.delta.thinking })}\n\n`,
              );
            }
          }
        } else if (evt.type === "result" && evt.result) {
          resultText = evt.result;
        }
      } catch {
        // non-JSON line — ignore
      }
    }
  });

  child.stderr.on("data", (data) => {
    console.error("[proxy] stderr:", data.toString().trim());
  });

  child.on("close", (code, signal) => {
    if (code !== 0 && !resultText) {
      res.write(
        `data: ${JSON.stringify({ type: "error", message: `claude CLI 退出 (code=${code}, signal=${signal})，请确认已登录：运行 claude 命令` })}\n\n`,
      );
    } else {
      res.write(
        `data: ${JSON.stringify({ type: "done", fullText: resultText.trim() })}\n\n`,
      );
    }
    clearInterval(keepalive);
    responseDone = true;
    res.end();
  });

  child.on("error", (err) => {
    console.error("[proxy] spawn error:", err);
    res.write(
      `data: ${JSON.stringify({ type: "error", message: `无法启动 claude CLI: ${err.message}` })}\n\n`,
    );
    clearInterval(keepalive);
    responseDone = true;
    res.end();
  });

  res.on("close", () => {
    clearInterval(keepalive);
    if (!responseDone && !child.killed) child.kill();
  });
});

if (existsSync(distPath)) {
  app.get("/{*path}", (req, res) => {
    res.sendFile(join(distPath, "index.html"));
  });
}

app.listen(PORT, "127.0.0.1", () => {
  console.log(`\n✅ WPS Claude PPT 代理服务器已启动`);
  console.log(`   地址: http://127.0.0.1:${PORT}`);
  console.log(`   健康检查: http://127.0.0.1:${PORT}/health`);
  if (existsSync(distPath)) {
    console.log(`   前端: http://127.0.0.1:${PORT}/ (dist 静态文件)`);
  }
  console.log(`   代码执行桥: /execute-code, /pending-code, /code-result`);
  console.log(`   ⚠️  PPT 插件独立端口，不影响 Excel 插件 (3001/5173)\n`);
  // #region agent log
  fetch('http://127.0.0.1:7244/ingest/63acb95d-6f91-4165-a07a-5bab2abb61eb',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'38bc8c'},body:JSON.stringify({sessionId:'38bc8c',location:'proxy-server.js:listen',message:'PPT proxy started',data:{port:PORT,pid:process.pid},timestamp:Date.now(),hypothesisId:'H1'})}).catch(()=>{});
  // #endregion
});
