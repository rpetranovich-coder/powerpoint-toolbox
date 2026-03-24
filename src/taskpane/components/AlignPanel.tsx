import React from "react";
import {
  makeStyles,
  Button,
  Tooltip,
  Divider,
} from "@fluentui/react-components";
import {
  AlignLeft20Regular,
  AlignCenterHorizontal20Regular,
  AlignRight20Regular,
  AlignTop20Regular,
  AlignCenterVertical20Regular,
  AlignBottom20Regular,
} from "@fluentui/react-icons";
import { alignShapes, distributeShapes, matchSize } from "../../lib/ppt";

interface AlignPanelProps {
  selectionCount: number;
  showToast: (msg: string, type: "success" | "error") => void;
}

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "4px" },
  row:  { display: "flex", flexWrap: "wrap", gap: "4px" },
  divider: { margin: "2px 0" },
});

export const AlignPanel: React.FC<AlignPanelProps> = ({ selectionCount, showToast }) => {
  const styles = useStyles();
  const can2 = selectionCount >= 2;
  const can3 = selectionCount >= 3;

  const run = async (fn: () => Promise<void>, msg: string) => {
    try { await fn(); showToast(msg, "success"); }
    catch (e: unknown) { showToast(e instanceof Error ? e.message : String(e), "error"); }
  };

  return (
    <div className={styles.root}>
      {/* Align edges */}
      <div className={styles.row}>
        <Tooltip content="Align Left" relationship="label">
          <Button icon={<AlignLeft20Regular />} disabled={!can2}
            onClick={() => run(() => alignShapes("left"), "Aligned left")} />
        </Tooltip>
        <Tooltip content="Center Horizontally" relationship="label">
          <Button icon={<AlignCenterHorizontal20Regular />} disabled={!can2}
            onClick={() => run(() => alignShapes("centerH"), "Centered horizontally")} />
        </Tooltip>
        <Tooltip content="Align Right" relationship="label">
          <Button icon={<AlignRight20Regular />} disabled={!can2}
            onClick={() => run(() => alignShapes("right"), "Aligned right")} />
        </Tooltip>
        <Tooltip content="Align Top" relationship="label">
          <Button icon={<AlignTop20Regular />} disabled={!can2}
            onClick={() => run(() => alignShapes("top"), "Aligned top")} />
        </Tooltip>
        <Tooltip content="Center Vertically" relationship="label">
          <Button icon={<AlignCenterVertical20Regular />} disabled={!can2}
            onClick={() => run(() => alignShapes("centerV"), "Centered vertically")} />
        </Tooltip>
        <Tooltip content="Align Bottom" relationship="label">
          <Button icon={<AlignBottom20Regular />} disabled={!can2}
            onClick={() => run(() => alignShapes("bottom"), "Aligned bottom")} />
        </Tooltip>
      </div>

      <Divider className={styles.divider} />

      {/* Distribute + Match size on one row */}
      <div className={styles.row}>
        <Tooltip content="Distribute Horizontally" relationship="label">
          <Button disabled={!can3}
            onClick={() => run(() => distributeShapes("horizontal"), "Distributed horizontally")}>
            ⟷ Horiz
          </Button>
        </Tooltip>
        <Tooltip content="Distribute Vertically" relationship="label">
          <Button disabled={!can3}
            onClick={() => run(() => distributeShapes("vertical"), "Distributed vertically")}>
            ↕ Vert
          </Button>
        </Tooltip>
        <Tooltip content="Match Width" relationship="label">
          <Button disabled={!can2}
            onClick={() => run(() => matchSize("width"), "Widths matched")}>
            ↔ W
          </Button>
        </Tooltip>
        <Tooltip content="Match Height" relationship="label">
          <Button disabled={!can2}
            onClick={() => run(() => matchSize("height"), "Heights matched")}>
            ↕ H
          </Button>
        </Tooltip>
        <Tooltip content="Match Width & Height" relationship="label">
          <Button disabled={!can2}
            onClick={() => run(() => matchSize("both"), "Sizes equalized")}>
            ⊞ Both
          </Button>
        </Tooltip>
      </div>
    </div>
  );
};
