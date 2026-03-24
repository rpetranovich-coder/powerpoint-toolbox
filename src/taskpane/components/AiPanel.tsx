import React, { useState } from "react";
import {
  makeStyles,
  Button,
  Input,
  Textarea,
  Spinner,
  tokens,
} from "@fluentui/react-components";
import {
  getSlideTextContent,
  insertActionTitle,
  getSlideShapeTexts,
  updateShapeTexts,
  ShapeTextEntry,
} from "../../lib/ppt";

interface AiPanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const LS_KEY = "tbx_claude_api_key";

const useStyles = makeStyles({
  root:       { display: "flex", flexDirection: "column", gap: "6px" },
  row:        { display: "flex", gap: "4px", alignItems: "center", flexWrap: "wrap" },
  keyInput:   { flex: 1 },
  taglineBox: { width: "100%" },
  keyToggle:  { fontSize: "10px", color: tokens.colorNeutralForeground3, cursor: "pointer", userSelect: "none" },
});

// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function callClaude(apiKey: string, prompt: string, maxTokens: number): Promise<string> {
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: maxTokens,
      messages: [{ role: "user", content: prompt }],
    }),
  });

  if (!resp.ok) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const err: any = await resp.json().catch(() => ({}));
    throw new Error(err?.error?.message ?? `API error ${resp.status}`);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const data: any = await resp.json();
  const text: string = data?.content?.[0]?.text?.trim() ?? "";
  if (!text) throw new Error("Empty response from Claude");
  return text;
}

export const AiPanel: React.FC<AiPanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [apiKey, setApiKey]         = useState(() => localStorage.getItem(LS_KEY) ?? "");
  const [keyVisible, setKeyVisible] = useState(() => !localStorage.getItem(LS_KEY));
  const [tagline, setTagline]       = useState("");
  const [generating, setGenerating] = useState(false);
  const [sharpening, setSharpening] = useState(false);

  const saveKey = (k: string) => {
    setApiKey(k);
    if (k) { localStorage.setItem(LS_KEY, k); setKeyVisible(false); }
    else   localStorage.removeItem(LS_KEY);
  };

  const generate = async () => {
    if (!apiKey.trim()) { showToast("Enter your Anthropic API key first", "error"); return; }
    setGenerating(true);
    try {
      const slideText = await getSlideTextContent();
      if (!slideText.trim()) { showToast("No text found on this slide", "error"); return; }

      const text = await callClaude(
        apiKey.trim(),
        `You are a McKinsey slide editor. Generate a single action title for the following slide content. ` +
        `The title must be a direct, insight-driven declarative statement (not a question or a topic label) ` +
        `of 10–18 words that captures the single most important takeaway a reader should remember. ` +
        `Return ONLY the title text — no quotes, no labels, no extra punctuation.\n\nSlide content:\n${slideText}`,
        120,
      );
      setTagline(text);
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    } finally {
      setGenerating(false);
    }
  };

  const insert = async () => {
    if (!tagline.trim()) { showToast("Generate a tagline first", "error"); return; }
    try {
      await insertActionTitle(tagline.trim());
      showToast("Action title inserted", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  const sharpen = async () => {
    if (!apiKey.trim()) { showToast("Enter your Anthropic API key first", "error"); return; }
    setSharpening(true);
    try {
      const shapes = await getSlideShapeTexts();
      if (shapes.length === 0) { showToast("No text found on this slide", "error"); return; }

      // Send as indexed JSON so Claude can return in exact same order
      const input = shapes.map((s, i) => ({ i, text: s.text }));
      const raw = await callClaude(
        apiKey.trim(),
        `You are an elite strategy consultant editor at Bain. ` +
        `Rewrite each text snippet below to be crisp, punchy, and action-oriented — ` +
        `the standard expected from a top internal strategy team. ` +
        `Rules: fix grammar, eliminate passive voice, cut filler words, lead with the insight, ` +
        `preserve all key data and numbers, keep bullet structure if present, ` +
        `do not merge separate bullets into one. ` +
        `Return a JSON array with the same length and order: [{"i":0,"text":"..."},{"i":1,"text":"..."},...]. ` +
        `Return ONLY the JSON array — no markdown, no explanation.\n\n` +
        `Input:\n${JSON.stringify(input)}`,
        2048,
      );

      // Parse — strip possible markdown fences
      const jsonStr = raw.replace(/^```[^\n]*\n?/, "").replace(/```$/, "").trim();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const parsed: { i: number; text: string }[] = JSON.parse(jsonStr);

      const updates: ShapeTextEntry[] = parsed
        .filter((r) => r.i >= 0 && r.i < shapes.length && typeof r.text === "string")
        .map((r) => ({ id: shapes[r.i].id, text: r.text }));

      await updateShapeTexts(updates);
      showToast(`${updates.length} text block${updates.length !== 1 ? "s" : ""} sharpened`, "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    } finally {
      setSharpening(false);
    }
  };

  const busy = generating || sharpening;

  return (
    <div className={styles.root}>
      {keyVisible ? (
        <div className={styles.row}>
          <Input
            className={styles.keyInput}
            type="password"
            placeholder="Anthropic API key"
            value={apiKey}
            onChange={(_, d) => saveKey(d.value)}
          />
        </div>
      ) : (
        <span className={styles.keyToggle} onClick={() => setKeyVisible(true)}>
          {apiKey ? "API key saved — click to change" : "Set API key"}
        </span>
      )}
      <div className={styles.row}>
        <Button
          onClick={generate}
          disabled={busy}
          icon={generating ? <Spinner size="tiny" /> : undefined}
        >
          {generating ? "Generating…" : "Generate Tagline"}
        </Button>
        <Button
          onClick={sharpen}
          disabled={busy}
          icon={sharpening ? <Spinner size="tiny" /> : undefined}
        >
          {sharpening ? "Sharpening…" : "Sharpen"}
        </Button>
      </div>
      {tagline && (
        <>
          <Textarea
            className={styles.taglineBox}
            value={tagline}
            onChange={(_, d) => setTagline(d.value)}
            rows={3}
          />
          <div className={styles.row}>
            <Button appearance="primary" onClick={insert}>
              Insert on Slide
            </Button>
          </div>
        </>
      )}
    </div>
  );
};
