import React, { useState } from "react";
import {
  makeStyles,
  Button,
  Text,
  ToggleButton,
  Tooltip,
  tokens,
  Divider,
} from "@fluentui/react-components";
import {
  AlignLeft20Regular,
  AlignCenterHorizontal20Regular,
  AlignRight20Regular,
  AlignTop20Regular,
  AlignCenterVertical20Regular,
  AlignBottom20Regular,
  ArrowLeft20Regular,
  ArrowRight20Regular,
  ArrowUp20Regular,
  ArrowDown20Regular,
} from "@fluentui/react-icons";
import {
  alignShapes,
  distributeShapes,
  matchSize,
  nudgeShapes,
} from "../../lib/ppt";

interface AlignPanelProps {
  selectionCount: number;
  showToast: (msg: string, type: "success" | "error") => void;
}

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "6px" },
  row: { display: "flex", flexWrap: "wrap", gap: "4px" },
  hint: { color: tokens.colorNeutralForeground3, fontSize: "11px" },
  divider: { margin: "4px 0" },
  label: { fontSize: "11px", color: tokens.colorNeutralForeground2, marginBottom: "2px" },
});

const NUDGE_AMOUNTS = [1, 5] as const;

export const AlignPanel: React.FC<AlignPanelProps> = ({
  selectionCount,
  showToast,
}) => {
  const styles = useStyles();
  const [nudgeAmt, setNudgeAmt] = useState<1 | 5>(1);

  const can2 = selectionCount >= 2;
  const can3 = selectionCount >= 3;
  const can1 = selectionCount >= 1;

  const run = async (fn: () => Promise<void>, successMsg: string) => {
    try {
      await fn();
      showToast(successMsg, "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  return (
    <div className={styles.root}>
      {!can2 && (
        <Text className={styles.hint}>Select 2+ shapes to align.</Text>
      )}

      {/* Edge alignment */}
      <Text className={styles.label}>Align edges</Text>
      <div className={styles.row}>
        <Tooltip content="Align Left" relationship="label">
          <Button
            size="small"
            icon={<AlignLeft20Regular />}
            disabled={!can2}
            onClick={() => run(() => alignShapes("left"), "Aligned left")}
          />
        </Tooltip>
        <Tooltip content="Center Horizontally" relationship="label">
          <Button
            size="small"
            icon={<AlignCenterHorizontal20Regular />}
            disabled={!can2}
            onClick={() => run(() => alignShapes("centerH"), "Centered horizontally")}
          />
        </Tooltip>
        <Tooltip content="Align Right" relationship="label">
          <Button
            size="small"
            icon={<AlignRight20Regular />}
            disabled={!can2}
            onClick={() => run(() => alignShapes("right"), "Aligned right")}
          />
        </Tooltip>
        <Tooltip content="Align Top" relationship="label">
          <Button
            size="small"
            icon={<AlignTop20Regular />}
            disabled={!can2}
            onClick={() => run(() => alignShapes("top"), "Aligned top")}
          />
        </Tooltip>
        <Tooltip content="Center Vertically" relationship="label">
          <Button
            size="small"
            icon={<AlignCenterVertical20Regular />}
            disabled={!can2}
            onClick={() => run(() => alignShapes("centerV"), "Centered vertically")}
          />
        </Tooltip>
        <Tooltip content="Align Bottom" relationship="label">
          <Button
            size="small"
            icon={<AlignBottom20Regular />}
            disabled={!can2}
            onClick={() => run(() => alignShapes("bottom"), "Aligned bottom")}
          />
        </Tooltip>
      </div>

      <Divider className={styles.divider} />

      {/* Distribution */}
      <Text className={styles.label}>Distribute (edge-based, 3+ shapes)</Text>
      <div className={styles.row}>
        <Button
          size="small"
          disabled={!can3}
          onClick={() =>
            run(
              () => distributeShapes("horizontal"),
              "Distributed horizontally"
            )
          }
        >
          ⟷ Horiz
        </Button>
        <Button
          size="small"
          disabled={!can3}
          onClick={() =>
            run(() => distributeShapes("vertical"), "Distributed vertically")
          }
        >
          ↕ Vert
        </Button>
      </div>

      <Divider className={styles.divider} />

      {/* Size matching */}
      <Text className={styles.label}>Match size (anchor = 1st selected)</Text>
      <div className={styles.row}>
        <Button
          size="small"
          disabled={!can2}
          onClick={() => run(() => matchSize("width"), "Widths matched")}
        >
          ↔ Width
        </Button>
        <Button
          size="small"
          disabled={!can2}
          onClick={() => run(() => matchSize("height"), "Heights matched")}
        >
          ↕ Height
        </Button>
        <Button
          size="small"
          disabled={!can2}
          onClick={() => run(() => matchSize("both"), "Sizes equalized")}
        >
          ⊞ Both
        </Button>
      </div>

      <Divider className={styles.divider} />

      {/* Nudge */}
      <Text className={styles.label}>Nudge</Text>
      <div className={styles.row}>
        {NUDGE_AMOUNTS.map((amt) => (
          <ToggleButton
            key={amt}
            size="small"
            checked={nudgeAmt === amt}
            onClick={() => setNudgeAmt(amt)}
          >
            {amt}pt
          </ToggleButton>
        ))}
      </div>
      <div className={styles.row}>
        <Button
          size="small"
          icon={<ArrowLeft20Regular />}
          disabled={!can1}
          onClick={() =>
            run(() => nudgeShapes("left", nudgeAmt), `Nudged left ${nudgeAmt}pt`)
          }
        />
        <Button
          size="small"
          icon={<ArrowRight20Regular />}
          disabled={!can1}
          onClick={() =>
            run(() => nudgeShapes("right", nudgeAmt), `Nudged right ${nudgeAmt}pt`)
          }
        />
        <Button
          size="small"
          icon={<ArrowUp20Regular />}
          disabled={!can1}
          onClick={() =>
            run(() => nudgeShapes("up", nudgeAmt), `Nudged up ${nudgeAmt}pt`)
          }
        />
        <Button
          size="small"
          icon={<ArrowDown20Regular />}
          disabled={!can1}
          onClick={() =>
            run(() => nudgeShapes("down", nudgeAmt), `Nudged down ${nudgeAmt}pt`)
          }
        />
      </div>
    </div>
  );
};
