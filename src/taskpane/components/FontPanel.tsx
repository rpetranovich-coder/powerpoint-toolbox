import React, { useState } from "react";
import { makeStyles, Button, Input, tokens } from "@fluentui/react-components";
import { standardizeFont } from "../../lib/ppt";

interface FontPanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const useStyles = makeStyles({
  root: { display: "flex", gap: "4px", alignItems: "center" },
  input: { flex: 1 },
});

export const FontPanel: React.FC<FontPanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [fontName, setFontName] = useState("Aptos");

  const handleApply = async () => {
    const name = fontName.trim();
    if (!name) return;
    try {
      await standardizeFont(name);
      showToast(`Font set to ${name} on this slide`, "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  return (
    <div className={styles.root}>
      <Input
        className={styles.input}
        size="small"
        value={fontName}
        placeholder="Font name"
        onChange={(_, d) => setFontName(d.value)}
        onKeyDown={(e) => { if (e.key === "Enter") handleApply(); }}
      />
      <Button size="small" disabled={!fontName.trim()} onClick={handleApply}>
        Apply to Slide
      </Button>
    </div>
  );
};
