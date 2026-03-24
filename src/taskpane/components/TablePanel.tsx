import React, { useState } from "react";
import { makeStyles, Button, Input, tokens } from "@fluentui/react-components";
import { insertTable } from "../../lib/ppt";

interface TablePanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const useStyles = makeStyles({
  row:      { display: "flex", gap: "4px", alignItems: "center" },
  numInput: { width: "44px" },
  muted:    { fontSize: "10px", color: tokens.colorNeutralForeground3 },
});

export const TablePanel: React.FC<TablePanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [rows, setRows] = useState(4);
  const [cols, setCols] = useState(3);

  const clamp = (v: number, min: number, max: number) =>
    Math.max(min, Math.min(max, v));

  const handleInsert = async () => {
    try {
      await insertTable(rows, cols);
      showToast(`${rows}×${cols} table inserted`, "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  return (
    <div className={styles.row}>
      <span className={styles.muted}>R</span>
      <Input
        className={styles.numInput}
               type="text" inputMode="numeric"
        value={String(rows)}
        onChange={(_, d) => setRows(clamp(Number(d.value) || 1, 1, 20))}
      />
      <span className={styles.muted}>C</span>
      <Input
        className={styles.numInput}
               type="text" inputMode="numeric"
        value={String(cols)}
        onChange={(_, d) => setCols(clamp(Number(d.value) || 1, 1, 10))}
      />
      <Button onClick={handleInsert}>
        Insert Table
      </Button>
    </div>
  );
};
