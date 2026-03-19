import React, { useState } from "react";
import { makeStyles, Button, Text, tokens } from "@fluentui/react-components";
import { Note20Regular, PaintBrush20Regular } from "@fluentui/react-icons";
import {
  insertStickyComment,
  readCommentTemplate,
  StickyCommentTemplate,
  DEFAULT_COMMENT_TEMPLATE,
} from "../../lib/ppt";

interface CommentPanelProps {
  selectedCommentName: string | null;
  showToast: (msg: string, type: "success" | "error") => void;
}

const STORAGE_KEY = "tbx_comment_template_v1";

function loadTemplate(): StickyCommentTemplate {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return { ...DEFAULT_COMMENT_TEMPLATE, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_COMMENT_TEMPLATE };
}

function saveTemplate(t: StickyCommentTemplate): void {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(t)); } catch { /* ignore */ }
}

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "8px" },
  row:  { display: "flex", flexWrap: "wrap", gap: "4px" },
  hint: { color: tokens.colorNeutralForeground3, fontSize: "11px", lineHeight: "1.45" },
  preview: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    fontSize: "10px",
    color: tokens.colorNeutralForeground3,
    paddingTop: "2px",
  },
  swatch: {
    display: "inline-block",
    width: "12px",
    height: "12px",
    borderRadius: "2px",
    flexShrink: 0,
  },
});

export const CommentPanel: React.FC<CommentPanelProps> = ({
  selectedCommentName,
  showToast,
}) => {
  const styles = useStyles();
  const [template, setTemplate] = useState<StickyCommentTemplate>(loadTemplate);
  const isCommentSelected = selectedCommentName !== null;

  const handleInsert = async () => {
    try {
      await insertStickyComment(template);
      showToast("Sticky comment inserted", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  const handleCapture = async () => {
    if (!selectedCommentName) return;
    try {
      const captured = await readCommentTemplate(selectedCommentName);
      setTemplate(captured);
      saveTemplate(captured);
      showToast("Default format updated from selected comment", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  return (
    <div className={styles.root}>
      <Text className={styles.hint}>
        {isCommentSelected
          ? "Sticky comment selected — use its look as the new default, or insert another."
          : "Inserts a new sticky comment using the current default format."}
      </Text>

      <div className={styles.row}>
        <Button size="small" icon={<Note20Regular />} onClick={handleInsert}>
          Add Sticky Comment
        </Button>
        {isCommentSelected && (
          <Button
            size="small"
            icon={<PaintBrush20Regular />}
            appearance="subtle"
            onClick={handleCapture}
          >
            Set as Default Format
          </Button>
        )}
      </div>

      {/* Format preview strip */}
      <div className={styles.preview}>
        <span
          className={styles.swatch}
          style={{
            background: `#${template.fillColor}`,
            border: `1px solid #${template.borderColor}`,
          }}
        />
        <span>
          {template.fontName} · {template.fontSize}pt
          {template.fontBold ? " · Bold" : ""}
          {template.fontItalic ? " · Italic" : ""}
          {" · "}{template.width}×{template.height}pt
        </span>
      </div>
    </div>
  );
};
