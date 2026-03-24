import React, { useState } from "react";
import {
  makeStyles,
  Button,
  Input,
  Label,
  Select,
} from "@fluentui/react-components";
import { insertStatusLabel } from "../../lib/ppt";

interface StatusPanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const PRESETS = ["/PRELIMINARY", "/DRAFT", "/CONFIDENTIAL", "/INTERNAL", "(custom)"];

const useStyles = makeStyles({
  root:  { display: "flex", flexDirection: "column", gap: "2px" },
  row:   { display: "flex", gap: "4px", alignItems: "flex-end" },
  input: { flex: 1 },
});

export const StatusPanel: React.FC<StatusPanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [preset, setPreset]         = useState(PRESETS[0]);
  const [customText, setCustomText] = useState("");

  const isCustom      = preset === "(custom)";
  const effectiveText = isCustom ? customText.trim() : preset;

  const handleInsert = async () => {
    if (!effectiveText) return;
    try { await insertStatusLabel(effectiveText); showToast("Status label inserted/updated", "success"); }
    catch (e: unknown) { showToast(e instanceof Error ? e.message : String(e), "error"); }
  };

  return (
    <div className={styles.root}>
      <Label>Status label</Label>
      <div className={styles.row}>
        {isCustom
          ? <Input className={styles.input} placeholder="/YOUR-LABEL"
              value={customText} onChange={(_, d) => setCustomText(d.value)} />
          : <Select className={styles.input} value={preset} onChange={(_, d) => setPreset(d.value)}>
              {PRESETS.map((p) => <option key={p} value={p}>{p}</option>)}
            </Select>
        }
        <Button disabled={!effectiveText} onClick={handleInsert}>
          Insert
        </Button>
      </div>
      {isCustom && (
        <Button appearance="subtle" onClick={() => setPreset(PRESETS[0])}>
          ← Back to presets
        </Button>
      )}
    </div>
  );
};
