import React, { useState } from "react";
import { makeStyles, Button, ToggleButton } from "@fluentui/react-components";
import { insertColumnGuides, clearColumnGuides } from "../../lib/ppt";

interface GuidesPanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const COL_OPTIONS = [2, 3, 4, 5] as const;
type ColOption = typeof COL_OPTIONS[number];

const useStyles = makeStyles({
  row: { display: "flex", gap: "4px", alignItems: "center", flexWrap: "wrap" },
});

export const GuidesPanel: React.FC<GuidesPanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [guideCols, setGuideCols] = useState<ColOption>(3);

  const run = async (fn: () => Promise<void>, msg: string) => {
    try { await fn(); showToast(msg, "success"); }
    catch (e: unknown) { showToast(e instanceof Error ? e.message : String(e), "error"); }
  };

  return (
    <div className={styles.row}>
      {COL_OPTIONS.map((n) => (
        <ToggleButton
          key={n}
          checked={guideCols === n}
          appearance={guideCols === n ? "primary" : "secondary"}
          onClick={() => setGuideCols(n)}
        >
          {n}
        </ToggleButton>
      ))}
      <Button
               onClick={() => run(() => insertColumnGuides(guideCols), `${guideCols}-column guides inserted`)}
      >
        Guides
      </Button>
      <Button
               appearance="subtle"
        onClick={() => run(clearColumnGuides, "Guides cleared")}
      >
        Clear
      </Button>
    </div>
  );
};
