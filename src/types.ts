export type MessageRole = "user" | "assistant" | "system";

export type InteractionMode = "agent" | "plan" | "ask";

export interface ModeEnforcement {
  codeBridge?: boolean | string;
  codeBlockRender?: boolean | string;
  maxTurns?: number;
  autoExecute?: boolean | string;
  stripCodeBlocks?: boolean | string;
  planUI?: boolean | string;
}

export interface ModeDefinition {
  id: string;
  name: string;
  description: string;
  default?: boolean;
  enforcement: ModeEnforcement;
  quickActions?: QuickAction[];
}

export interface ChatMessage {
  id: string;
  role: MessageRole;
  content: string;
  timestamp: number;
  codeBlocks?: CodeBlock[];
  isStreaming?: boolean;
  isError?: boolean;
  thinkingMs?: number;
  thinkingContent?: string;
  suggestAgentSwitch?: boolean;
}

export interface CodeBlock {
  id: string;
  language: string;
  code: string;
  executed?: boolean;
  result?: string;
  error?: string;
}

export interface ShapeInfo {
  id: number;
  name: string;
  type: string;
  text?: string;
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SlideInfo {
  index: number;
  layout: string;
  shapeCount: number;
  texts: string[];
}

export interface PptContext {
  fileName: string;
  slideCount: number;
  currentSlideIndex: number | null;
  pageSetup: { width: number; height: number };
  selectedShapes: ShapeInfo[];
  slideTexts: SlideInfo[];
  timestamp?: number;
}

export interface ModelOption {
  id: string;
  label: string;
  description: string;
  cliModel: string;
}

export const MODEL_OPTIONS: ModelOption[] = [
  {
    id: "sonnet",
    label: "Sonnet 4.6",
    description: "最佳编程模型，速度与质量兼顾",
    cliModel: "claude-sonnet-4-6",
  },
  {
    id: "opus",
    label: "Opus 4.6",
    description: "最强推理能力，适合复杂分析",
    cliModel: "claude-opus-4-6",
  },
  {
    id: "haiku",
    label: "Haiku 4.5",
    description: "极速响应，轻量任务首选",
    cliModel: "claude-haiku-4-5",
  },
];

export interface AttachmentFile {
  name: string;
  content: string;
  size: number;
  type?: "text" | "image" | "table";
  tempPath?: string;
  previewUrl?: string;
}

export interface FixErrorRequest {
  code: string;
  error: string;
  language: string;
}

export interface QuickAction {
  icon: string;
  label: string;
  prompt: string;
}

export interface AddToChatPayload {
  type: string;
  slideIndex: number;
  shapeNames: string[];
  texts: string[];
  timestamp: number;
}
