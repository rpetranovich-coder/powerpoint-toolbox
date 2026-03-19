import React, { useState } from "react";
import {
  makeStyles,
  Button,
  Input,
  Label,
  Divider,
} from "@fluentui/react-components";
import { insertFootnote, insertSource } from "../../lib/ppt";

interface FootnotePanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const useStyles = makeStyles({
  root:   { display: "flex", flexDirection: "column", gap: "6px" },
  field:  { display: "flex", flexDirection: "column", gap: "2px" },
  row:    { display: "flex", gap: "4px", alignItems: "flex-end" },
  input:  { flex: 1 },
  divider: { margin: "2px 0" },
});

export const FootnotePanel: React.FC<FootnotePanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [footnoteText, setFootnoteText] = useState("");
  const [sourceText, setSourceText]     = useState("");

  const run = async (fn: () => Promise<void>, msg: string) => {
    try { await fn(); showToast(msg, "success"); }
    catch (e: unknown) { showToast(e instanceof Error ? e.message : String(e), "error"); }
  };

  return (
    <div className={styles.root}>
      <div className={styles.field}>
        <Label size="small">Footnote</Label>
        <div className={styles.row}>
          <Input className={styles.input} size="small" placeholder="e.g. Based on 2024 data"
            value={footnoteText} onChange={(_, d) => setFootnoteText(d.value)} />
          <Button size="small" disabled={!footnoteText.trim()}
            onClick={() => run(() => insertFootnote(footnoteText.trim()), "Footnote inserted/updated")}>
            Insert
          </Button>
        </div>
      </div>

      <Divider className={styles.divider} />

      <div className={styles.field}>
        <Label size="small">Source</Label>
        <div className={styles.row}>
          <Input className={styles.input} size="small" placeholder="e.g. McKinsey Global Institute"
            value={sourceText} onChange={(_, d) => setSourceText(d.value)} />
          <Button size="small" disabled={!sourceText.trim()}
            onClick={() => run(() => insertSource(sourceText.trim()), "Source inserted/updated")}>
            Insert
          </Button>
        </div>
      </div>
    </div>
  );
};
