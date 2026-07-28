import React, { useState } from "react";
import {
  makeStyles,
  Button,
  Dropdown,
  Option,
  tokens,
} from "@fluentui/react-components";
import {
  TextBulletListLtr20Regular,
  ArrowSync20Regular,
} from "@fluentui/react-icons";
import {
  insertBulletsBox,
  resizeBulletsBox,
  syncSubBullets,
  setBulletsSpacing,
} from "../../lib/ppt";
import { stepFontSize } from "../../lib/fontSizes";

interface BulletsPanelProps {
  selectedBulletsName: string | null;
  showToast: (msg: string, type: "success" | "error") => void;
}

const SPACING_LEVELS = [
  { label: "Compressed", before: 0,  after: 0  },
  { label: "Tight",      before: 2,  after: 2  },
  { label: "Normal",     before: 6,  after: 6  },
  { label: "Wide",       before: 12, after: 12 },
  { label: "Ultra-wide", before: 20, after: 20 },
] as const;
type SpacingLabel = typeof SPACING_LEVELS[number]["label"];
const DEFAULT_SPACING_LABEL: SpacingLabel = "Normal";
const DEFAULT_L1_SIZE = 16;

const useStyles = makeStyles({
  root:        { display: "flex", flexDirection: "column", gap: "6px" },
  row:         { display: "flex", gap: "4px", alignItems: "center", flexWrap: "wrap" },
  sizeDisplay: { fontSize: "12px", minWidth: "32px", textAlign: "center" },
  label:       { fontSize: "11px", color: tokens.colorNeutralForeground3 },
  dropdown:    { minWidth: "110px" },
});

export const BulletsPanel: React.FC<BulletsPanelProps> = ({
  selectedBulletsName,
  showToast,
}) => {
  const styles = useStyles();
  const [l1Size, setL1Size] = useState<number>(DEFAULT_L1_SIZE);
  const [spacingLabel, setSpacingLabel] = useState<SpacingLabel>(DEFAULT_SPACING_LABEL);

  const spacing = SPACING_LEVELS.find((s) => s.label === spacingLabel) ?? SPACING_LEVELS[2];
  const hasSelection = selectedBulletsName !== null;

  const run = async (fn: () => Promise<void>, msg: string) => {
    try { await fn(); showToast(msg, "success"); }
    catch (e: unknown) { showToast(e instanceof Error ? e.message : String(e), "error"); }
  };

  const handleInsert = () =>
    run(
      () => insertBulletsBox(l1Size, { before: spacing.before, after: spacing.after }),
      "Bullets inserted"
    );

  const stepSize = async (delta: number) => {
    const next = stepFontSize(l1Size, delta);
    if (next === l1Size) return;
    setL1Size(next);
    if (selectedBulletsName) {
      await run(
        () => resizeBulletsBox(selectedBulletsName, next),
        `Bullets resized to ${next}pt`
      );
    }
  };

  const handleSpacingChange = async (label: SpacingLabel) => {
    setSpacingLabel(label);
    const next = SPACING_LEVELS.find((s) => s.label === label);
    if (!next || !selectedBulletsName) return;
    await run(
      () => setBulletsSpacing(selectedBulletsName, next.before, next.after),
      `Spacing set to ${label.toLowerCase()}`
    );
  };

  const handleSync = () => {
    if (!selectedBulletsName) return;
    return run(
      () => syncSubBullets(selectedBulletsName),
      "Sub-bullets resynced"
    );
  };

  return (
    <div className={styles.root}>
      <div className={styles.row}>
        <Button
          size="small"
          icon={<TextBulletListLtr20Regular />}
          onClick={handleInsert}
        >
          Insert Bullets
        </Button>
      </div>

      <div className={styles.row}>
        <span className={styles.label}>L1</span>
        <Button size="small" onClick={() => stepSize(-1)}>−</Button>
        <span className={styles.sizeDisplay}>{l1Size}pt</span>
        <Button size="small" onClick={() => stepSize(+1)}>+</Button>
        <Button
          size="small"
          appearance="subtle"
          icon={<ArrowSync20Regular />}
          disabled={!hasSelection}
          onClick={handleSync}
        >
          Sync
        </Button>
      </div>

      <div className={styles.row}>
        <span className={styles.label}>Spacing</span>
        <Dropdown
          className={styles.dropdown}
          size="small"
          value={spacingLabel}
          selectedOptions={[spacingLabel]}
          onOptionSelect={(_, data) => {
            const v = data.optionValue as SpacingLabel | undefined;
            if (v) handleSpacingChange(v);
          }}
        >
          {SPACING_LEVELS.map((s) => (
            <Option key={s.label} value={s.label}>{s.label}</Option>
          ))}
        </Dropdown>
      </div>
    </div>
  );
};
