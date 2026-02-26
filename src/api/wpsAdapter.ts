/**
 * WPS PPT 数据适配层
 *
 * 架构：Plugin Host main.js 定时将 PPT 数据 POST 到 proxy-server，
 * 本模块通过 GET /wps-context 获取最新数据。
 * 代码执行：POST /execute-code 提交 → 轮询 /code-result/:id 获取结果。
 */
import type { PptContext, AddToChatPayload } from "../types";

const PROXY_URL = "http://127.0.0.1:3002";

let _wpsAvailable = false;

export function isWpsAvailable(): boolean {
  return _wpsAvailable;
}

export async function getPptContext(): Promise<PptContext> {
  try {
    const res = await fetch(`${PROXY_URL}/wps-context`, {
      signal: AbortSignal.timeout(1500),
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();

    if (data.error || (!data.fileName && !data.slideCount)) {
      _wpsAvailable = false;
      return getMockContext();
    }

    _wpsAvailable = true;
    return {
      fileName: data.fileName ?? "",
      slideCount: data.slideCount ?? 0,
      currentSlideIndex: data.currentSlideIndex ?? null,
      pageSetup: data.pageSetup ?? { width: 960, height: 540 },
      selectedShapes: data.selectedShapes ?? [],
      slideTexts: data.slideTexts ?? [],
    };
  } catch {
    _wpsAvailable = false;
    return getMockContext();
  }
}

const CODE_RESULT_POLL_MS = 300;
const CODE_RESULT_TIMEOUT_MS = 30000;

export interface ExecuteResult {
  result: string;
}

export async function executeCode(code: string): Promise<ExecuteResult> {
  const submitRes = await fetch(`${PROXY_URL}/execute-code`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ code }),
  });

  if (!submitRes.ok) {
    throw new Error(`提交代码失败: HTTP ${submitRes.status}`);
  }

  const { id } = await submitRes.json();
  const deadline = Date.now() + CODE_RESULT_TIMEOUT_MS;

  while (Date.now() < deadline) {
    await sleep(CODE_RESULT_POLL_MS);

    const pollRes = await fetch(`${PROXY_URL}/code-result/${id}`);
    if (!pollRes.ok) continue;

    const data = await pollRes.json();
    if (!data.ready) continue;

    if (data.error) {
      throw new Error(data.error);
    }
    return { result: data.result ?? "执行成功" };
  }

  throw new Error("代码执行超时（30秒）");
}

export async function pollAddToChat(): Promise<AddToChatPayload | null> {
  try {
    const res = await fetch(`${PROXY_URL}/add-to-chat/poll`);
    if (!res.ok) return null;
    const data = await res.json();
    if (!data.pending) return null;
    return data as AddToChatPayload;
  } catch {
    return null;
  }
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

let _lastCtxJson = "";

export function onContextChange(
  callback: (ctx: PptContext) => void,
): () => void {
  let active = true;
  const POLL_INTERVAL = 2500;

  const poll = async () => {
    if (!active) return;
    try {
      const ctx = await getPptContext();
      const json = JSON.stringify({
        fileName: ctx.fileName,
        slideCount: ctx.slideCount,
        currentSlideIndex: ctx.currentSlideIndex,
        selectedShapeCount: ctx.selectedShapes.length,
      });

      if (json !== _lastCtxJson) {
        _lastCtxJson = json;
        callback(ctx);
      }
    } catch {
      // ignore
    }
    if (active) setTimeout(poll, POLL_INTERVAL);
  };

  setTimeout(poll, POLL_INTERVAL);
  return () => {
    active = false;
  };
}

function getMockContext(): PptContext {
  return {
    fileName: "示例演示.pptx",
    slideCount: 5,
    currentSlideIndex: 1,
    pageSetup: { width: 960, height: 540 },
    selectedShapes: [
      {
        id: 1,
        name: "标题 1",
        type: "Placeholder",
        text: "Q1 季度汇报",
        left: 50,
        top: 30,
        width: 860,
        height: 80,
      },
    ],
    slideTexts: [
      {
        index: 1,
        layout: "Title Slide",
        shapeCount: 2,
        texts: ["Q1 季度汇报", "2026年1月-3月"],
      },
      {
        index: 2,
        layout: "Title and Content",
        shapeCount: 3,
        texts: ["业绩概览", "总收入: 1500万", "同比增长 25%"],
      },
      {
        index: 3,
        layout: "Two Content",
        shapeCount: 4,
        texts: ["产品线分析", "产品A: 800万", "产品B: 500万", "产品C: 200万"],
      },
      {
        index: 4,
        layout: "Title and Content",
        shapeCount: 2,
        texts: ["市场展望", "预计 Q2 增长 30%"],
      },
      {
        index: 5,
        layout: "Title Only",
        shapeCount: 1,
        texts: ["谢谢"],
      },
    ],
  };
}
