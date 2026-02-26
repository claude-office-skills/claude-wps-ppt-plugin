import { useState, useEffect, useRef, memo, useCallback } from "react";
import type { QuickAction, InteractionMode } from "../types";
import styles from "./QuickActionCards.module.css";

const PROXY_BASE = "http://127.0.0.1:3002";
const DEBOUNCE_MS = 1500;
const VISIBLE_COUNT = 4;

interface CommandDef {
  id: string;
  icon: string;
  label: string;
  description: string;
  scope: string;
  prompt: string;
}

interface ModeDef {
  id: string;
  quickActions?: Array<{
    icon: string;
    label: string;
    prompt: string;
    scope?: string;
  }>;
}

const FALLBACK_GENERAL: QuickAction[] = [
  { icon: "📝", label: "生成幻灯片", prompt: "帮我生成一组完整的幻灯片" },
  { icon: "✨", label: "美化布局", prompt: "帮我美化当前演示文稿的布局" },
  { icon: "🎤", label: "添加备注", prompt: "帮我为每页幻灯片生成演讲者备注" },
  { icon: "📄", label: "提取文本", prompt: "提取当前演示文稿中所有文本内容" },
];

const FALLBACK_SELECTION: QuickAction[] = [
  { icon: "🔧", label: "修改形状", prompt: "帮我修改当前选中的形状" },
  { icon: "🌐", label: "翻译文本", prompt: "翻译选中形状中的文本为英文" },
  { icon: "🎨", label: "调整样式", prompt: "调整选中形状的视觉样式" },
  { icon: "🖼️", label: "插入图片", prompt: "在当前页插入图片" },
];

function toQuickAction(cmd: CommandDef): QuickAction {
  return { icon: cmd.icon, label: cmd.label, prompt: cmd.prompt };
}

interface Props {
  hasSelection: boolean;
  onAction: (prompt: string) => void;
  disabled?: boolean;
  mode?: InteractionMode;
}

const QuickActionCards = memo(function QuickActionCards({
  hasSelection,
  onAction,
  disabled,
  mode = "agent",
}: Props) {
  const [generalCmds, setGeneralCmds] =
    useState<QuickAction[]>(FALLBACK_GENERAL);
  const [selectionCmds, setSelectionCmds] =
    useState<QuickAction[]>(FALLBACK_SELECTION);
  const [modeActions, setModeActions] = useState<
    Record<string, { general: QuickAction[]; selection: QuickAction[] }>
  >({});
  const [stableHasSelection, setStableHasSelection] = useState(hasSelection);
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    if (hasSelection === stableHasSelection) {
      if (timerRef.current) {
        clearTimeout(timerRef.current);
        timerRef.current = null;
      }
      return;
    }
    timerRef.current = setTimeout(() => {
      setStableHasSelection(hasSelection);
      timerRef.current = null;
    }, DEBOUNCE_MS);
    return () => {
      if (timerRef.current) clearTimeout(timerRef.current);
    };
  }, [hasSelection, stableHasSelection]);

  useEffect(() => {
    fetch(`${PROXY_BASE}/commands`)
      .then((r) => r.json())
      .then((cmds: CommandDef[]) => {
        const gen = cmds
          .filter((c) => c.scope === "general")
          .map(toQuickAction);
        const sel = cmds
          .filter((c) => c.scope === "selection")
          .map(toQuickAction);
        if (gen.length > 0) setGeneralCmds(gen);
        if (sel.length > 0) setSelectionCmds(sel);
      })
      .catch(() => {});

    fetch(`${PROXY_BASE}/modes`)
      .then((r) => r.json())
      .then((modes: ModeDef[]) => {
        const result: Record<
          string,
          { general: QuickAction[]; selection: QuickAction[] }
        > = {};
        for (const m of modes) {
          if (m.quickActions && m.quickActions.length > 0) {
            result[m.id] = {
              general: m.quickActions
                .filter((a) => !a.scope || a.scope === "general")
                .map((a) => ({
                  icon: a.icon,
                  label: a.label,
                  prompt: a.prompt,
                })),
              selection: m.quickActions
                .filter((a) => a.scope === "selection")
                .map((a) => ({
                  icon: a.icon,
                  label: a.label,
                  prompt: a.prompt,
                })),
            };
          }
        }
        setModeActions(result);
      })
      .catch(() => {});
  }, []);

  const modeSpecific = modeActions[mode];

  const allActions: QuickAction[] = [];
  const seen = new Set<string>();
  const addUnique = (list: QuickAction[]) => {
    for (const a of list) {
      if (!seen.has(a.label)) {
        seen.add(a.label);
        allActions.push(a);
      }
    }
  };

  if (stableHasSelection) {
    if (modeSpecific?.selection.length) addUnique(modeSpecific.selection);
    addUnique(selectionCmds);
    if (modeSpecific?.general.length) addUnique(modeSpecific.general);
    addUnique(generalCmds);
  } else {
    if (modeSpecific?.general.length) addUnique(modeSpecific.general);
    addUnique(generalCmds);
    if (modeSpecific?.selection.length) addUnique(modeSpecific.selection);
    addUnique(selectionCmds);
  }

  const actions = allActions;

  const [expanded, setExpanded] = useState(false);
  const toggleExpand = useCallback(() => setExpanded((v) => !v), []);

  const hasMore = actions.length > VISIBLE_COUNT;
  const visibleActions = expanded ? actions : actions.slice(0, VISIBLE_COUNT);

  return (
    <div className={`${styles.grid} ${expanded ? styles.gridExpanded : ""}`}>
      {visibleActions.map((action) => (
        <button
          key={action.label}
          className={styles.card}
          onClick={() => onAction(action.prompt)}
          disabled={disabled}
        >
          <span className={styles.cardIcon}>{action.icon}</span>
          <span className={styles.cardLabel}>{action.label}</span>
        </button>
      ))}
      {hasMore && (
        <button
          className={`${styles.card} ${styles.moreBtn}`}
          onClick={toggleExpand}
        >
          <span className={styles.cardIcon}>{expanded ? "‹" : "›"}</span>
          <span className={styles.cardLabel}>
            {expanded ? "收起" : `更多+${actions.length - VISIBLE_COUNT}`}
          </span>
        </button>
      )}
    </div>
  );
});

export default QuickActionCards;
