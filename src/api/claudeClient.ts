/**
 * Claude API 调用层（PPT 版）
 *
 * 通过本地代理服务器（proxy-server.js）调用 claude CLI，
 * 绕过 OAuth token 无法直接调用 Anthropic 公开 API 的限制。
 *
 * 使用 XMLHttpRequest + onprogress 实现流式接收，
 * 因为 WPS CEF 浏览器的 fetch ReadableStream 不支持增量读取。
 */
import type { PptContext, ChatMessage, AttachmentFile } from "../types";

const PROXY_BASE = "http://127.0.0.1:3002";

function buildContextString(ctx: PptContext): string {
  let s = `演示文稿: ${ctx.fileName}\n`;
  s += `总页数: ${ctx.slideCount}\n`;
  s += `页面尺寸: ${ctx.pageSetup.width} x ${ctx.pageSetup.height}\n`;

  if (ctx.currentSlideIndex) {
    s += `\n⚠️ 当前页: 第 ${ctx.currentSlideIndex} 页 — 请优先操作此页，除非用户另有要求。\n`;
  }

  if (ctx.selectedShapes.length > 0) {
    s += `\n[选中形状] 共 ${ctx.selectedShapes.length} 个:\n`;
    for (const shape of ctx.selectedShapes) {
      s += `  - ${shape.name} (${shape.type}) 位置(${shape.left},${shape.top}) 尺寸(${shape.width}x${shape.height})`;
      if (shape.text) {
        const preview = shape.text.length > 100 ? shape.text.substring(0, 100) + "..." : shape.text;
        s += ` 文本: "${preview}"`;
      }
      s += "\n";
    }
  }

  if (ctx.slideTexts.length > 0) {
    s += `\n[各页内容摘要]\n`;
    for (const slide of ctx.slideTexts) {
      const textsPreview = slide.texts.join(" | ");
      const truncated = textsPreview.length > 200 ? textsPreview.substring(0, 200) + "..." : textsPreview;
      s += `  第${slide.index}页 (${slide.layout || "默认布局"}, ${slide.shapeCount}个形状): ${truncated}\n`;
    }
    if (ctx.slideCount > 20) {
      s += `  ... (共 ${ctx.slideCount} 页，仅展示前 20 页)\n`;
    }
  }

  return s;
}

export interface StreamCallbacks {
  onToken: (token: string) => void;
  onThinking?: (text: string) => void;
  onComplete: (fullText: string) => void;
  onError: (err: Error) => void;
  onModeInfo?: (mode: string, enforcement: Record<string, unknown>) => void;
}

export async function checkProxy(): Promise<boolean> {
  try {
    const resp = await fetch(`${PROXY_BASE}/health`, {
      signal: AbortSignal.timeout(2000),
    });
    return resp.ok;
  } catch {
    return false;
  }
}

export interface SendMessageOptions {
  model?: string;
  attachments?: AttachmentFile[];
  signal?: AbortSignal;
  webSearch?: boolean;
  mode?: string;
}

/**
 * SSE line parser — processes complete lines from buffer, dispatches callbacks.
 * Returns the remaining (possibly incomplete) tail of the buffer.
 */
function processSseLines(
  buffer: string,
  fullTextRef: { v: string },
  callbacks: StreamCallbacks,
): string {
  const lines = buffer.split("\n");
  const tail = lines.pop() ?? "";

  for (const line of lines) {
    if (!line.startsWith("data: ")) continue;
    const data = line.slice(6).trim();
    try {
      const event = JSON.parse(data);
      if (event.type === "mode") {
        callbacks.onModeInfo?.(event.mode, event.enforcement);
      } else if (event.type === "token") {
        fullTextRef.v += event.text;
        callbacks.onToken(event.text);
      } else if (event.type === "thinking") {
        callbacks.onThinking?.(event.text);
      } else if (event.type === "done") {
        if (event.fullText && !fullTextRef.v) fullTextRef.v = event.fullText;
      } else if (event.type === "error") {
        throw new Error(event.message);
      }
    } catch (parseErr) {
      if (
        parseErr instanceof Error &&
        parseErr.message !== "Unexpected token"
      ) {
        throw parseErr;
      }
    }
  }
  return tail;
}

/** 流式发送消息给 Claude — 使用 XHR onprogress 实现流式 */
export async function sendMessage(
  userMessage: string,
  history: ChatMessage[],
  pptCtx: PptContext,
  callbacks: StreamCallbacks,
  options?: SendMessageOptions,
): Promise<void> {
  const messages = [
    ...history
      .filter((m) => m.role === "user" || m.role === "assistant")
      .map((m) => ({
        role: m.role as "user" | "assistant",
        content: m.content,
      })),
    { role: "user" as const, content: userMessage },
  ];

  const context = buildContextString(pptCtx);

  const payload: Record<string, unknown> = { messages, context };
  if (options?.model) payload.model = options.model;
  if (options?.mode) payload.mode = options.mode;
  if (options?.webSearch) payload.webSearch = true;
  if (options?.attachments?.length) {
    payload.attachments = options.attachments.map((f) => ({
      name: f.name,
      content: f.content,
      type: f.type ?? "text",
      tempPath: f.tempPath,
    }));
  }

  const fullTextRef = { v: "" };
  const streamStart = Date.now();
  let progressCount = 0;
  let prevLen = 0;

  return new Promise<void>((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", `${PROXY_BASE}/chat`, true);
    xhr.setRequestHeader("Content-Type", "application/json");

    let buffer = "";
    let aborted = false;

    if (options?.signal) {
      options.signal.addEventListener("abort", () => {
        aborted = true;
        xhr.abort();
      });
    }

    xhr.onprogress = () => {
      const newData = xhr.responseText.slice(prevLen);
      prevLen = xhr.responseText.length;
      if (!newData) return;

      progressCount++;
      buffer += newData;

      // #region agent log
      if (progressCount <= 5 || progressCount % 20 === 0) { fetch('http://127.0.0.1:7244/ingest/63acb95d-6f91-4165-a07a-5bab2abb61eb',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'38bc8c'},body:JSON.stringify({sessionId:'38bc8c',location:'ppt-claudeClient.ts:xhr-progress',message:'PPT XHR onprogress fired',data:{progressCount,newDataLen:newData.length,totalLen:prevLen,elapsed:Date.now()-streamStart,fullTextLen:fullTextRef.v.length},timestamp:Date.now(),hypothesisId:'H-stream-fix'})}).catch(()=>{}); }
      // #endregion

      try {
        buffer = processSseLines(buffer, fullTextRef, callbacks);
      } catch (err) {
        aborted = true;
        xhr.abort();
        reject(err instanceof Error ? err : new Error(String(err)));
      }
    };

    xhr.onload = () => {
      if (aborted) return;
      if (xhr.status !== 200) {
        reject(new Error(`代理服务器错误 ${xhr.status}: ${xhr.responseText}`));
        return;
      }
      const remaining = xhr.responseText.slice(prevLen);
      if (remaining) {
        buffer += remaining;
        try {
          processSseLines(buffer, fullTextRef, callbacks);
        } catch (err) {
          reject(err instanceof Error ? err : new Error(String(err)));
          return;
        }
      }
      // #region agent log
      fetch('http://127.0.0.1:7244/ingest/63acb95d-6f91-4165-a07a-5bab2abb61eb',{method:'POST',headers:{'Content-Type':'application/json','X-Debug-Session-Id':'38bc8c'},body:JSON.stringify({sessionId:'38bc8c',location:'ppt-claudeClient.ts:xhr-done',message:'PPT XHR stream completed',data:{totalProgressEvents:progressCount,totalElapsed:Date.now()-streamStart,fullTextLen:fullTextRef.v.length},timestamp:Date.now(),hypothesisId:'H-stream-fix'})}).catch(()=>{});
      // #endregion
      callbacks.onComplete(fullTextRef.v);
      resolve();
    };

    xhr.onerror = () => {
      if (aborted) {
        callbacks.onComplete(fullTextRef.v || "（已中止生成）");
        resolve();
        return;
      }
      reject(new Error("无法连接代理服务器，请检查 proxy-server 是否运行"));
    };

    xhr.onabort = () => {
      callbacks.onComplete(fullTextRef.v || "（已中止生成）");
      resolve();
    };

    xhr.send(JSON.stringify(payload));
  });
}

export function extractCodeBlocks(
  text: string,
): Array<{ language: string; code: string }> {
  const regex = /```(\w+)?\n([\s\S]*?)```/g;
  const blocks: Array<{ language: string; code: string }> = [];
  let match;
  while ((match = regex.exec(text)) !== null) {
    blocks.push({ language: match[1] || "javascript", code: match[2].trim() });
  }
  return blocks;
}
