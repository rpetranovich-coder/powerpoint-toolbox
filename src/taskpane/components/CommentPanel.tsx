import React, { useState } from "react";
import { makeStyles, Button } from "@fluentui/react-components";
import { Note20Regular, PaintBrush20Regular, Circle20Regular } from "@fluentui/react-icons";
import {
  insertStickyComment,
  readCommentTemplate,
  StickyCommentTemplate,
  DEFAULT_COMMENT_TEMPLATE,
  insertStoplight,
  readStoplightTemplate,
  StoplightTemplate,
  DEFAULT_STOPLIGHT_TEMPLATE,
} from "../../lib/ppt";

interface CommentPanelProps {
  selectedCommentName: string | null;
  selectedStoplightName: string | null;
  showToast: (msg: string, type: "success" | "error") => void;
}

const COMMENT_STORAGE_KEY   = "tbx_comment_template_v1";
const STOPLIGHT_STORAGE_KEY = "tbx_stoplight_template_v1";

function loadCommentTemplate(): StickyCommentTemplate {
  try {
    const raw = localStorage.getItem(COMMENT_STORAGE_KEY);
    if (raw) return { ...DEFAULT_COMMENT_TEMPLATE, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_COMMENT_TEMPLATE };
}

function saveCommentTemplate(t: StickyCommentTemplate): void {
  try { localStorage.setItem(COMMENT_STORAGE_KEY, JSON.stringify(t)); } catch { /* ignore */ }
}

function loadStoplightTemplate(): StoplightTemplate {
  try {
    const raw = localStorage.getItem(STOPLIGHT_STORAGE_KEY);
    if (raw) return { ...DEFAULT_STOPLIGHT_TEMPLATE, ...JSON.parse(raw) };
  } catch { /* ignore */ }
  return { ...DEFAULT_STOPLIGHT_TEMPLATE };
}

function saveStoplightTemplate(t: StoplightTemplate): void {
  try { localStorage.setItem(STOPLIGHT_STORAGE_KEY, JSON.stringify(t)); } catch { /* ignore */ }
}

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "8px" },
  row:  { display: "flex", flexWrap: "wrap", gap: "4px" },
});

export const CommentPanel: React.FC<CommentPanelProps> = ({
  selectedCommentName,
  selectedStoplightName,
  showToast,
}) => {
  const styles = useStyles();
  const [commentTemplate, setCommentTemplate] = useState<StickyCommentTemplate>(loadCommentTemplate);
  const [stoplightTemplate, setStoplightTemplate] = useState<StoplightTemplate>(loadStoplightTemplate);

  const handleInsertComment = async () => {
    try {
      await insertStickyComment(commentTemplate);
      showToast("Sticky comment inserted", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  const handleCaptureComment = async () => {
    if (!selectedCommentName) return;
    try {
      const captured = await readCommentTemplate(selectedCommentName);
      setCommentTemplate(captured);
      saveCommentTemplate(captured);
      showToast("Default comment format updated", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  const handleInsertStoplight = async () => {
    try {
      await insertStoplight(stoplightTemplate);
      showToast("Stoplight inserted", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  const handleCaptureStoplight = async () => {
    if (!selectedStoplightName) return;
    try {
      const captured = await readStoplightTemplate(selectedStoplightName);
      setStoplightTemplate(captured);
      saveStoplightTemplate(captured);
      showToast("Default stoplight format updated", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.row}>
        <Button size="small" icon={<Note20Regular />} onClick={handleInsertComment}>
          Add Sticky Comment
        </Button>
        {selectedCommentName && (
          <Button
            icon={<PaintBrush20Regular />}
            appearance="subtle"
            size="small"
            onClick={handleCaptureComment}
          >
            Set as Default Format
          </Button>
        )}
      </div>
      <div className={styles.row}>
        <Button size="small" icon={<Circle20Regular />} onClick={handleInsertStoplight}>
          Add Stoplight
        </Button>
        {selectedStoplightName && (
          <Button
            icon={<PaintBrush20Regular />}
            appearance="subtle"
            size="small"
            onClick={handleCaptureStoplight}
          >
            Set as Default Format
          </Button>
        )}
      </div>
    </div>
  );
};
