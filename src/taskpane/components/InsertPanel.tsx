import React, { useState } from "react";
import {
  makeStyles,
  Button,
  Input,
  ToggleButton,
  Divider,
  tokens,
} from "@fluentui/react-components";
import { insertTable, insertColumnGuides, clearColumnGuides } from "../../lib/ppt";

interface InsertPanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const COL_OPTIONS = [2, 3, 4, 5] as const;
type ColOption = typeof COL_OPTIONS[number];

const useStyles = makeStyles({
  root:     { display: "flex", flexDirection: "column", gap: "4px" },
  row:      { display: "flex", gap: "4px", alignItems: "center", flexWrap: "wrap" },
  numInput: { width: "44px" },
  divider:  { margin: "2px 0" },
  muted:    { fontSize: "10px", color: tokens.colorNeutralForeground3 },
});

export const InsertPanel: React.FC<InsertPanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [rows, setRows]         = useState(4);
  const [cols, setCols]         = useState(3);
  const [guideCols, setGuideCols] = useState<ColOption>(3);

  const clamp = (v: number, min: number, max: number) =>
    Math.max(min, Math.min(max, v));

  const run = async (fn: () => Promise<void>, msg: string) => {
    try { await fn(); showToast(msg, "success"); }
    catch (e: unknown) { showToast(e instanceof Error ? e.message : String(e), "error"); }
  };

  return (
    <div className={styles.root}>

      {/* ── Table ── */}
      <div className={styles.row}>
        <span className={styles.muted}>R</span>
        <Input
          className={styles.numInput}
          size="small"
          type="number"
          value={String(rows)}
          onChange={(_, d) => setRows(clamp(Number(d.value) || 1, 1, 20))}
        />
        <span className={styles.muted}>C</span>
        <Input
          className={styles.numInput}
          size="small"
          type="number"
          value={String(cols)}
          onChange={(_, d) => setCols(clamp(Number(d.value) || 1, 1, 10))}
        />
        <Button
          size="small"
          onClick={() => run(() => insertTable(rows, cols), `${rows}×${cols} table inserted`)}
        >
          Insert Table
        </Button>
      </div>

      <Divider className={styles.divider} />

      {/* ── Column guides ── */}
      <div className={styles.row}>
        {COL_OPTIONS.map((n) => (
          <ToggleButton
            key={n}
            size="small"
            checked={guideCols === n}
            onClick={() => setGuideCols(n)}
          >
            {n}
          </ToggleButton>
        ))}
        <Button
          size="small"
          onClick={() =>
            run(() => insertColumnGuides(guideCols), `${guideCols}-column guides inserted`)
          }
        >
          Guides
        </Button>
        <Button
          size="small"
          appearance="subtle"
          onClick={() => run(clearColumnGuides, "Guides cleared")}
        >
          Clear
        </Button>
      </div>

    </div>
  );
};
