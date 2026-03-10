import React, { useState } from "react";
import {
  makeStyles,
  Button,
  Input,
  Label,
  Select,
  Text,
  tokens,
} from "@fluentui/react-components";
import { insertStatusLabel } from "../../lib/ppt";

interface StatusPanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const PRESETS = [
  "/PRELIMINARY",
  "/DRAFT",
  "/CONFIDENTIAL",
  "/INTERNAL",
  "(custom)",
];

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "8px" },
  row: { display: "flex", gap: "4px", alignItems: "flex-end" },
  field: { display: "flex", flexDirection: "column", gap: "2px" },
  input: { flex: 1 },
  hint: { color: tokens.colorNeutralForeground3, fontSize: "11px" },
});

export const StatusPanel: React.FC<StatusPanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [preset, setPreset] = useState(PRESETS[0]);
  const [customText, setCustomText] = useState("");

  const isCustom = preset === "(custom)";
  const effectiveText = isCustom ? customText.trim() : preset;

  const handleInsert = async () => {
    if (!effectiveText) return;
    try {
      await insertStatusLabel(effectiveText);
      showToast("Status label inserted/updated", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  return (
    <div className={styles.root}>
      <Text className={styles.hint}>
        Inserts at top-right corner. Updates existing{" "}
        <code>TBX_STATUS</code> shape in place.
      </Text>

      <div className={styles.field}>
        <Label size="small">Preset</Label>
        <Select
          size="small"
          value={preset}
          onChange={(_, d) => setPreset(d.value)}
        >
          {PRESETS.map((p) => (
            <option key={p} value={p}>
              {p}
            </option>
          ))}
        </Select>
      </div>

      {isCustom && (
        <div className={styles.field}>
          <Label size="small">Custom label</Label>
          <Input
            size="small"
            placeholder="/YOUR-LABEL"
            value={customText}
            onChange={(_, d) => setCustomText(d.value)}
          />
        </div>
      )}

      <Button
        size="small"
        disabled={!effectiveText}
        onClick={handleInsert}
      >
        Insert / Update Status
      </Button>
    </div>
  );
};
