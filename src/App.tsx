import { useState, useRef, useEffect, useCallback } from "react";
import { nanoid } from "nanoid";
import { Claude } from "@lobehub/icons";
import MessageBubble from "./components/MessageBubble";
import ModelSelector from "./components/ModelSelector";
import ModeSelector from "./components/ModeSelector";
import AttachmentMenu from "./components/AttachmentMenu";
import QuickActionCards from "./components/QuickActionCards";
import HistoryPanel from "./components/HistoryPanel";
import { sendMessage, extractCodeBlocks, checkProxy } from "./api/claudeClient";
import {
  getPptContext,
  onContextChange,
  isWpsAvailable,
  executeCode,
  pollAddToChat,
} from "./api/wpsAdapter";
import {
  saveSession,
  loadSession,
  listSessions,
  generateTitle,
} from "./api/sessionStore";
import type {
  ChatMessage,
  PptContext,
  CodeBlock,
  AttachmentFile,
  InteractionMode,
} from "./types";
import styles from "./App.module.css";

const WELCOME_MESSAGE: ChatMessage = {
  id: "welcome",
  role: "assistant",
  content:
    "你好！我是 Claude，你的 WPS PPT AI 助手。\n\n我可以帮你：\n- **生成幻灯片**（大纲、内容、批量创建）\n- **美化布局**（对齐、配色、字体统一）\n- **处理文本**（提取、翻译、摘要）\n- **添加备注**（演讲者笔记、时间估算）\n\n请告诉我你想做什么，或者选中一个形状后提问。",
  timestamp: Date.now(),
};

export default function App() {
  const [messages, setMessages] = useState<ChatMessage[]>([WELCOME_MESSAGE]);
  const [input, setInput] = useState("");
  const [pptCtx, setPptCtx] = useState<PptContext | null>(null);
  const [loading, setLoading] = useState(false);
  const [proxyMissing, setProxyMissing] = useState(false);
  const [applyingMsgId, setApplyingMsgId] = useState<string | null>(null);
  const [selectedModel, setSelectedModel] = useState("claude-sonnet-4-6");
  const [attachedFiles, setAttachedFiles] = useState<AttachmentFile[]>([]);
  const [webSearchEnabled, setWebSearchEnabled] = useState(false);

  const [sessionId, setSessionId] = useState<string>(nanoid());
  const [historyOpen, setHistoryOpen] = useState(false);

  const [inputBoxHeight, setInputBoxHeight] = useState(100);
  const [currentMode, setCurrentMode] = useState<InteractionMode>(
    () =>
      (localStorage.getItem("wps-claude-mode") as InteractionMode) || "agent",
  );

  const abortRef = useRef<AbortController | null>(null);
  const lastSentInputRef = useRef<string>("");
  const bottomRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const dragStartRef = useRef<{ y: number; h: number } | null>(null);

  useEffect(() => {
    const initCtx = async () => {
      const ctx = await getPptContext();
      setPptCtx(ctx);
    };
    initCtx();
    const unsubscribe = onContextChange((ctx) => setPptCtx(ctx));
    return unsubscribe;
  }, []);

  useEffect(() => {
    let timer: ReturnType<typeof setTimeout> | null = null;
    let stopped = false;
    const poll = async () => {
      if (stopped) return;
      const ok = await checkProxy();
      setProxyMissing(!ok);
      timer = setTimeout(poll, ok ? 30000 : 3000);
    };
    poll();
    return () => {
      stopped = true;
      if (timer) clearTimeout(timer);
    };
  }, []);

  useEffect(() => {
    let stopped = false;
    const poll = async () => {
      if (stopped) return;
      try {
        const data = await pollAddToChat();
        if (data) {
          const label = `第${data.slideIndex}页 [${data.shapeNames.join(", ")}]`;
          const preview = data.texts.join("\n");
          setInput((prev) =>
            prev
              ? `${prev}\n\n📎 ${label}\n${preview}`
              : `📎 ${label}\n${preview}`,
          );
        }
      } catch { /* ignore */ }
      if (!stopped) setTimeout(poll, 1000);
    };
    poll();
    return () => { stopped = true; };
  }, []);

  useEffect(() => {
    const restoreLastSession = async () => {
      try {
        const sessions = await listSessions();
        if (sessions.length === 0) return;
        const latest = sessions[0];
        const session = await loadSession(latest.id);
        if (!session || !session.messages || session.messages.length === 0)
          return;
        setSessionId(session.id);
        setMessages([WELCOME_MESSAGE, ...session.messages]);
        if (session.model) setSelectedModel(session.model);
      } catch { /* ignore */ }
    };
    restoreLastSession();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    const realMessages = messages.filter((m) => m.id !== "welcome");
    if (realMessages.length === 0) return;
    const hasStreaming = realMessages.some((m) => m.isStreaming);
    if (hasStreaming) return;

    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      const title = generateTitle(realMessages);
      saveSession(sessionId, realMessages, { title, model: selectedModel });
    }, 1000);

    return () => {
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    };
  }, [messages, sessionId, selectedModel]);

  useEffect(() => {
    const handleGlobalCopy = (e: KeyboardEvent) => {
      if (!(e.metaKey || e.ctrlKey) || e.key !== "c") return;
      const active = document.activeElement;
      if (
        active instanceof HTMLTextAreaElement ||
        active instanceof HTMLInputElement
      )
        return;
      const sel = window.getSelection();
      const text = sel?.toString();
      if (!text) return;
      e.preventDefault();
      navigator.clipboard?.writeText(text).catch(() => {
        const ta = document.createElement("textarea");
        ta.value = text;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
      });
    };
    document.addEventListener("keydown", handleGlobalCopy);
    return () => document.removeEventListener("keydown", handleGlobalCopy);
  }, []);

  useEffect(() => {
    const isStreaming = messages.some((m) => m.isStreaming);
    if (isStreaming) {
      const container = bottomRef.current?.parentElement;
      if (container) container.scrollTop = container.scrollHeight;
    } else {
      bottomRef.current?.scrollIntoView({ behavior: "smooth" });
    }
  }, [messages]);

  const handleCodeExecuted = useCallback(
    (msgId: string, blockId: string, result: string, error?: string) => {
      setMessages((prev) =>
        prev.map((msg) => {
          if (msg.id !== msgId) return msg;
          const updatedBlocks = msg.codeBlocks?.map((b) =>
            b.id === blockId ? { ...b, executed: true, result, error } : b,
          );
          return { ...msg, codeBlocks: updatedBlocks };
        }),
      );
    },
    [],
  );

  const messagesRef = useRef(messages);
  messagesRef.current = messages;

  const handleApplyCode = useCallback(async (msgId: string) => {
    const msg = messagesRef.current.find((m) => m.id === msgId);
    const blocks = msg?.codeBlocks?.filter((b) => !b.executed);
    if (!blocks?.length) return;

    setApplyingMsgId(msgId);

    for (const block of blocks) {
      try {
        const { result } = await executeCode(block.code);
        setMessages((prev) =>
          prev.map((m) => {
            if (m.id !== msgId) return m;
            const updated = m.codeBlocks?.map((b) =>
              b.id === block.id ? { ...b, executed: true, result } : b,
            );
            return { ...m, codeBlocks: updated };
          }),
        );
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : String(err);
        setMessages((prev) =>
          prev.map((m) => {
            if (m.id !== msgId) return m;
            const updated = m.codeBlocks?.map((b) =>
              b.id === block.id ? { ...b, executed: true, error: errorMsg } : b,
            );
            return { ...m, codeBlocks: updated };
          }),
        );
        break;
      }
    }

    setApplyingMsgId(null);
  }, []);

  const loadingRef = useRef(loading);
  loadingRef.current = loading;

  const handleSendRef = useRef<(text?: string) => Promise<void>>(
    null as unknown as (text?: string) => Promise<void>,
  );

  const handleRetryFix = useCallback(
    (code: string, error: string, language: string) => {
      if (loadingRef.current) return;
      const fixPrompt = `代码执行出错，请修复以下所有错误并重新生成完整代码。

**错误信息：**
\`\`\`
${error}
\`\`\`

**原始代码（${language}）：**
\`\`\`${language}
${code}
\`\`\`

请修复所有问题，生成可直接执行的完整代码。注意：
1. 使用 WPS PPT 兼容的 API
2. 使用 Application.ActivePresentation 获取当前演示文稿
3. 对可能失败的操作添加 try/catch 保护
4. 使用 shape.TextFrame（不是 TextFrame2）`;
      handleSendRef.current(fixPrompt);
    },
    [],
  );

  const handleModeChange = useCallback((mode: InteractionMode) => {
    setCurrentMode(mode);
    localStorage.setItem("wps-claude-mode", mode);
  }, []);

  const handleFileAttach = useCallback((file: AttachmentFile) => {
    setAttachedFiles((prev) => [...prev, file]);
  }, []);

  const handleRemoveFile = useCallback((name: string) => {
    setAttachedFiles((prev) => prev.filter((f) => f.name !== name));
  }, []);

  const handleInputChange = useCallback(
    (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      setInput(e.target.value);
    },
    [],
  );

  const handleQuickAction = useCallback((prompt: string) => {
    handleSendRef.current(prompt);
  }, []);

  const handleToggleWebSearch = useCallback(() => {
    setWebSearchEnabled((v) => !v);
  }, []);

  const handleSendClick = useCallback(() => {
    handleSendRef.current();
  }, []);

  const handleOpenHistory = useCallback(() => {
    setHistoryOpen(true);
  }, []);

  const handleCloseHistory = useCallback(() => {
    setHistoryOpen(false);
  }, []);

  const handleNewChat = useCallback(() => {
    if (abortRef.current) {
      abortRef.current.abort();
      abortRef.current = null;
    }
    setSessionId(nanoid());
    setMessages([WELCOME_MESSAGE]);
    setLoading(false);
    setApplyingMsgId(null);
    setAttachedFiles([]);
  }, []);

  const handleSwitchToAgent = useCallback(() => {
    setCurrentMode("agent");
    localStorage.setItem("wps-claude-mode", "agent");
  }, []);

  const handleSend = async (text?: string) => {
    const userText = (text ?? input).trim();
    if (!userText || loading) return;

    const currentAttachments = [...attachedFiles];
    lastSentInputRef.current = userText;
    setInput("");
    setAttachedFiles([]);

    let displayContent = userText;
    if (currentAttachments.length > 0) {
      const otherAttachments = currentAttachments.filter(
        (f) => f.type !== "table",
      );
      if (otherAttachments.length > 0) {
        displayContent += `\n\n[附件: ${otherAttachments.map((f) => f.name).join(", ")}]`;
      }
    }

    const userMsg: ChatMessage = {
      id: nanoid(),
      role: "user",
      content: displayContent,
      timestamp: Date.now(),
    };

    const assistantMsgId = nanoid();
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: "assistant",
      content: "",
      timestamp: Date.now(),
      isStreaming: true,
    };

    setMessages((prev) => [...prev, userMsg, assistantMsg]);
    setLoading(true);

    const controller = new AbortController();
    abortRef.current = controller;

    let fullText = "";
    let thinkingText = "";
    const thinkingStart = Date.now();
    let firstTokenReceived = false;
    const modeSnapshot = currentMode;

    await sendMessage(
      userText,
      messages.filter((m) => m.id !== "welcome"),
      pptCtx ?? {
        fileName: "",
        slideCount: 0,
        currentSlideIndex: null,
        pageSetup: { width: 960, height: 540 },
        selectedShapes: [],
        slideTexts: [],
      },
      {
        onThinking: (text) => {
          thinkingText += text;
          setMessages((prev) =>
            prev.map((m) =>
              m.id === assistantMsgId
                ? { ...m, thinkingContent: thinkingText }
                : m,
            ),
          );
        },
        onToken: (token) => {
          fullText += token;
          const updates: Partial<ChatMessage> = { content: fullText };
          if (!firstTokenReceived) {
            firstTokenReceived = true;
            updates.thinkingMs = Date.now() - thinkingStart;
            setProxyMissing(false);
          }
          setMessages((prev) =>
            prev.map((m) =>
              m.id === assistantMsgId ? { ...m, ...updates } : m,
            ),
          );
        },
        onComplete: async (text) => {
          if (modeSnapshot === "ask") {
            const strippedText = text.replace(
              /```[\w]*\n[\s\S]*?```/g,
              "_(此处为代码操作，请切换至 Agent 模式执行)_",
            );

            const hadCode = strippedText !== text;
            const ACTION_HINTS =
              /切换.{0,4}Agent|switch.{0,6}agent|需要执行|需要操作|建议.{0,4}Agent/i;
            const suggestSwitch = hadCode || ACTION_HINTS.test(text);

            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantMsgId
                  ? {
                      ...m,
                      content: strippedText,
                      isStreaming: false,
                      codeBlocks: [],
                      suggestAgentSwitch: suggestSwitch,
                    }
                  : m,
              ),
            );
            setLoading(false);
            return;
          }

          const rawBlocks = extractCodeBlocks(text);
          const codeBlocks: CodeBlock[] = rawBlocks.map((b) => ({
            id: nanoid(),
            language: b.language,
            code: b.code,
          }));

          setMessages((prev) =>
            prev.map((m) =>
              m.id === assistantMsgId
                ? { ...m, content: text, isStreaming: false, codeBlocks }
                : m,
            ),
          );
          setLoading(false);

          const shouldAutoExecute = modeSnapshot === "agent";

          if (shouldAutoExecute && codeBlocks.length > 0) {
            setApplyingMsgId(assistantMsgId);
            for (let _bi = 0; _bi < codeBlocks.length; _bi++) {
              const block = codeBlocks[_bi];
              try {
                const { result } = await executeCode(block.code);
                setMessages((prev) =>
                  prev.map((m) => {
                    if (m.id !== assistantMsgId) return m;
                    const updated = m.codeBlocks?.map((b) =>
                      b.id === block.id ? { ...b, executed: true, result } : b,
                    );
                    return { ...m, codeBlocks: updated };
                  }),
                );
              } catch (err) {
                const errorMsg =
                  err instanceof Error ? err.message : String(err);
                setMessages((prev) =>
                  prev.map((m) => {
                    if (m.id !== assistantMsgId) return m;
                    const updated = m.codeBlocks?.map((b) =>
                      b.id === block.id
                        ? { ...b, executed: true, error: errorMsg }
                        : b,
                    );
                    return { ...m, codeBlocks: updated };
                  }),
                );
                break;
              }
            }
            setApplyingMsgId(null);
          }
        },
        onError: (err) => {
          const isProxyError =
            err.message.includes("fetch") ||
            err.message.includes("Failed") ||
            err.message.includes("代理");
          setMessages((prev) =>
            prev.map((m) =>
              m.id === assistantMsgId
                ? {
                    ...m,
                    content: isProxyError
                      ? "**错误**：无法连接代理服务器。\n\n请在终端运行：\n```\ncd ~/需求讨论/claude-wps-ppt-plugin\nnode proxy-server.js\n```"
                      : `**错误**：${err.message}`,
                    isStreaming: false,
                    isError: true,
                  }
                : m,
            ),
          );
          setProxyMissing(true);
          setLoading(false);
        },
      },
      {
        model: selectedModel,
        mode: currentMode,
        attachments:
          currentAttachments.length > 0 ? currentAttachments : undefined,
        signal: controller.signal,
        webSearch: webSearchEnabled,
      },
    );

    abortRef.current = null;
    lastSentInputRef.current = "";
  };

  handleSendRef.current = handleSend;

  const handleStop = () => {
    if (abortRef.current) {
      abortRef.current.abort();
      abortRef.current = null;
    }

    const savedInput = lastSentInputRef.current;
    lastSentInputRef.current = "";

    setMessages((prev) => {
      const streamingIdx = prev.findIndex((m) => m.isStreaming);
      if (streamingIdx === -1) return prev;
      const userMsgIdx = streamingIdx - 1;
      return prev.filter((_, i) => i !== streamingIdx && i !== userMsgIdx);
    });

    setInput(savedInput);
    setLoading(false);

    requestAnimationFrame(() => {
      textareaRef.current?.focus();
    });
  };

  const toBase64 = (buf: ArrayBuffer): string =>
    btoa(new Uint8Array(buf).reduce((d, b) => d + String.fromCharCode(b), ""));

  const handlePaste = async (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    const clipboardData = e.clipboardData;
    if (!clipboardData) return;

    const imageItem = Array.from(clipboardData.items).find((item) =>
      item.type.startsWith("image/"),
    );
    if (imageItem) {
      e.preventDefault();
      const file = imageItem.getAsFile();
      if (!file) return;
      try {
        const arrayBuf = await file.arrayBuffer();
        const base64 = toBase64(arrayBuf);
        const ext = file.type.split("/")[1] || "png";
        const fileName = `clipboard-${Date.now()}.${ext}`;
        const resp = await fetch("http://127.0.0.1:3002/upload-temp", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ base64, fileName }),
        });
        const result = await resp.json();
        if (result.ok) {
          const previewUrl = URL.createObjectURL(file);
          setAttachedFiles((prev) => [
            ...prev,
            {
              name: fileName,
              content: `[图片: ${fileName}]`,
              size: file.size,
              type: "image",
              tempPath: result.filePath,
              previewUrl,
            },
          ]);
        }
      } catch { /* ignore */ }
      return;
    }

    const plainText = clipboardData.getData("text/plain");
    if (plainText) {
      e.preventDefault();
      const ta = textareaRef.current;
      if (!ta) return;
      const start = ta.selectionStart;
      const end = ta.selectionEnd;
      setInput((prev) => prev.slice(0, start) + plainText + prev.slice(end));
      requestAnimationFrame(() => {
        const pos = start + plainText.length;
        if (textareaRef.current) {
          textareaRef.current.selectionStart = pos;
          textareaRef.current.selectionEnd = pos;
        }
      });
    }
  };

  const insertTextAtCursor = (pasteText: string) => {
    const ta = textareaRef.current;
    if (!ta) return;
    const start = ta.selectionStart;
    const end = ta.selectionEnd;
    setInput((prev) => prev.slice(0, start) + pasteText + prev.slice(end));
    requestAnimationFrame(() => {
      const pos = start + pasteText.length;
      if (textareaRef.current) {
        textareaRef.current.selectionStart = pos;
        textareaRef.current.selectionEnd = pos;
      }
    });
  };

  const fallbackNativePaste = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text) insertTextAtCursor(text);
    } catch { /* ignore */ }
  };

  const pasteViaProxy = async () => {
    try {
      const resp = await fetch("http://127.0.0.1:3002/clipboard");
      const data = await resp.json();
      if (!data.ok) {
        await fallbackNativePaste();
        return;
      }

      if (data.type === "image" && data.filePath) {
        const fileName = data.fileName || `clipboard-${Date.now()}.png`;
        setAttachedFiles((prev) => [
          ...prev,
          {
            name: fileName,
            content: `[图片: ${fileName}]`,
            size: 0,
            type: "image" as const,
            tempPath: data.filePath,
          },
        ]);
        return;
      }

      if (data.text) {
        insertTextAtCursor(data.text);
      }
    } catch {
      await fallbackNativePaste();
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }

    const isMod = e.metaKey || e.ctrlKey;
    if (!isMod) return;

    if (e.key === "v") {
      e.preventDefault();
      pasteViaProxy();
      return;
    }

    if (e.key === "a") {
      e.preventDefault();
      const ta = textareaRef.current;
      if (ta) {
        ta.selectionStart = 0;
        ta.selectionEnd = ta.value.length;
      }
    }

    if (e.key === "x") {
      e.preventDefault();
      const ta = textareaRef.current;
      if (!ta) return;
      const start = ta.selectionStart;
      const end = ta.selectionEnd;
      if (start === end) return;
      const selectedText = input.slice(start, end);
      navigator.clipboard?.writeText(selectedText).catch(() => {});
      setInput(input.slice(0, start) + input.slice(end));
      requestAnimationFrame(() => {
        ta.selectionStart = start;
        ta.selectionEnd = start;
      });
    }

    if (e.key === "c") {
      e.preventDefault();
      const ta = textareaRef.current;
      if (!ta) return;
      const start = ta.selectionStart;
      const end = ta.selectionEnd;
      if (start === end) return;
      navigator.clipboard?.writeText(input.slice(start, end)).catch(() => {});
    }
  };

  const inputBoxHeightRef = useRef(inputBoxHeight);
  inputBoxHeightRef.current = inputBoxHeight;

  const handleDragStart = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    dragStartRef.current = { y: e.clientY, h: inputBoxHeightRef.current };
    const onMove = (ev: MouseEvent) => {
      if (!dragStartRef.current) return;
      const delta = dragStartRef.current.y - ev.clientY;
      const next = Math.max(80, Math.min(400, dragStartRef.current.h + delta));
      setInputBoxHeight(next);
    };
    const onUp = () => {
      dragStartRef.current = null;
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }, []);

  const ctxLabel = pptCtx
    ? pptCtx.currentSlideIndex
      ? `第 ${pptCtx.currentSlideIndex}/${pptCtx.slideCount} 页` +
        (pptCtx.selectedShapes.length > 0
          ? ` · 已选 ${pptCtx.selectedShapes.length} 个形状`
          : "")
      : `${pptCtx.slideCount} 页幻灯片`
    : null;

  return (
    <div className={styles.shell}>
      <header className={styles.header}>
        <div className={styles.logoRow}>
          <div className={styles.logoIcon}>
            <Claude.Color size={20} />
          </div>
          <div className={styles.logoName}>Claude for PPT</div>
          <span className={styles.betaBadge}>Beta</span>
        </div>
        <div className={styles.headerActions}>
          <button
            className={styles.headerBtn}
            onClick={handleOpenHistory}
            title="历史记录"
          >
            <HistoryIcon />
          </button>
          <button
            className={styles.headerBtn}
            onClick={handleNewChat}
            title="新对话"
          >
            <NewChatIcon />
          </button>
        </div>
      </header>

      {ctxLabel && (
        <div className={styles.ctxBar}>
          <SlideIcon />
          <span className={styles.ctxText}>{ctxLabel}</span>
          <span className={styles.ctxBadge}>
            {isWpsAvailable() ? "WPS" : "mock"}
          </span>
        </div>
      )}

      {proxyMissing && (
        <div className={styles.warning}>
          ⚠ 代理服务器未运行，请在终端执行：node proxy-server.js
        </div>
      )}

      <div className={styles.chatArea}>
        {messages.map((msg) => (
          <MessageBubble
            key={msg.id}
            message={msg}
            onCodeExecuted={handleCodeExecuted}
            onApplyCode={handleApplyCode}
            onRetryFix={handleRetryFix}
            isApplying={applyingMsgId === msg.id}
            onSwitchToAgent={handleSwitchToAgent}
          />
        ))}
        <div ref={bottomRef} />
      </div>

      <div className={styles.inputArea}>
        <QuickActionCards
          hasSelection={!!(pptCtx?.selectedShapes?.length)}
          onAction={handleQuickAction}
          disabled={loading}
          mode={currentMode}
        />

        <div className={styles.inputBox} style={{ height: inputBoxHeight }}>
          <div className={styles.dragHandle} onMouseDown={handleDragStart}>
            <div className={styles.dragDots} />
          </div>

          <div className={styles.inputBody}>
            <div className={styles.inputChips}>
              {attachedFiles.map((f) => (
                <span
                  key={f.name}
                  className={`${styles.inlineChip} ${f.type === "image" ? styles.chipImage : ""}`}
                >
                  {f.type === "image" ? (
                    f.previewUrl ? (
                      <img
                        src={f.previewUrl}
                        alt={f.name}
                        className={styles.chipThumb}
                      />
                    ) : (
                      <ImageIcon />
                    )
                  ) : (
                    <span className={styles.chipFileIcon}>📎</span>
                  )}
                  <span className={styles.chipLabel}>{f.name}</span>
                  <button
                    className={styles.chipRemove}
                    onClick={() => handleRemoveFile(f.name)}
                  >
                    ×
                  </button>
                </span>
              ))}
            </div>

            <textarea
              ref={textareaRef}
              className={styles.inlineTextarea}
              value={input}
              onChange={handleInputChange}
              onKeyDown={handleKeyDown}
              onPaste={handlePaste}
              placeholder={
                attachedFiles.length > 0
                  ? "描述你想做什么..."
                  : "发个指令...（Enter 发送，Shift+Enter 换行）"
              }
              rows={1}
              disabled={loading}
            />
          </div>

          <div className={styles.inputToolbar}>
            <div className={styles.toolbarLeft}>
              <AttachmentMenu
                onFileAttach={handleFileAttach}
                webSearchEnabled={webSearchEnabled}
                onToggleWebSearch={handleToggleWebSearch}
                disabled={loading}
              />
              <ModeSelector
                mode={currentMode}
                onChange={handleModeChange}
                disabled={loading}
              />
            </div>
            <div className={styles.toolbarRight}>
              <ModelSelector
                value={selectedModel}
                onChange={setSelectedModel}
                disabled={loading}
              />
              {loading ? (
                <button
                  className={`${styles.sendBtn} ${styles.stopBtn}`}
                  onClick={handleStop}
                  title="停止生成"
                >
                  <StopIcon />
                </button>
              ) : (
                <button
                  className={styles.sendBtn}
                  onClick={handleSendClick}
                  disabled={!input.trim()}
                  title="发送"
                >
                  <SendIcon />
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

      <HistoryPanel
        visible={historyOpen}
        onClose={handleCloseHistory}
        currentSessionId={sessionId}
        onSelectSession={async (id) => {
          const session = await loadSession(id);
          if (!session) return;
          if (abortRef.current) {
            abortRef.current.abort();
            abortRef.current = null;
          }
          setSessionId(session.id);
          setMessages(
            session.messages.length > 0
              ? [WELCOME_MESSAGE, ...session.messages]
              : [WELCOME_MESSAGE],
          );
          if (session.model) setSelectedModel(session.model);
          setLoading(false);
          setApplyingMsgId(null);
        }}
      />
    </div>
  );
}

function HistoryIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="12" cy="12" r="10" />
      <polyline points="12 6 12 12 16 14" />
    </svg>
  );
}

function SlideIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
      <rect x="2" y="3" width="20" height="14" rx="2" />
      <path d="M8 21h8" />
      <path d="M12 17v4" />
    </svg>
  );
}

function SendIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" stroke="none">
      <path d="M3.478 2.404a.75.75 0 0 0-.926.941l2.432 7.905H13.5a.75.75 0 0 1 0 1.5H4.984l-2.432 7.905a.75.75 0 0 0 .926.94l18.04-8.01a.75.75 0 0 0 0-1.37L3.478 2.404Z" />
    </svg>
  );
}

function StopIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" stroke="none">
      <rect x="6" y="6" width="12" height="12" rx="2" />
    </svg>
  );
}

function NewChatIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z" />
      <line x1="9" y1="10" x2="15" y2="10" />
      <line x1="12" y1="7" x2="12" y2="13" />
    </svg>
  );
}

function ImageIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="3" width="18" height="18" rx="2" />
      <circle cx="8.5" cy="8.5" r="1.5" />
      <path d="M21 15l-5-5L5 21" />
    </svg>
  );
}
