import React from "react";
import { makeStyles, Button, Text, tokens, Divider } from "@fluentui/react-components";
import {
  groupShapes,
  ungroupShapes,
  setZOrder,
  ZOrderType,
} from "../../lib/ppt";

interface GroupPanelProps {
  selectionCount: number;
  showToast: (msg: string, type: "success" | "error") => void;
}

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "6px" },
  row: { display: "flex", flexWrap: "wrap", gap: "4px" },
  hint: { color: tokens.colorNeutralForeground3, fontSize: "11px" },
  label: { fontSize: "11px", color: tokens.colorNeutralForeground2, marginBottom: "2px" },
  divider: { margin: "4px 0" },
});

export const GroupPanel: React.FC<GroupPanelProps> = ({
  selectionCount,
  showToast,
}) => {
  const styles = useStyles();
  const can1 = selectionCount >= 1;
  const can2 = selectionCount >= 2;

  const run = async (fn: () => Promise<void>, successMsg: string) => {
    try {
      await fn();
      showToast(successMsg, "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  const zOrder = (type: ZOrderType, label: string) =>
    run(() => setZOrder(type), label);

  return (
    <div className={styles.root}>
      {!can2 && (
        <Text className={styles.hint}>Select 2+ shapes to group.</Text>
      )}

      <Text className={styles.label}>Group</Text>
      <div className={styles.row}>
        <Button
          size="small"
          disabled={!can2}
          onClick={() => run(groupShapes, "Grouped")}
        >
          Group
        </Button>
        <Button
          size="small"
          disabled={!can1}
          onClick={() => run(ungroupShapes, "Ungrouped")}
        >
          Ungroup
        </Button>
      </div>

      <Divider className={styles.divider} />

      <Text className={styles.label}>Z-Order</Text>
      <div className={styles.row}>
        <Button
          size="small"
          disabled={!can1}
          onClick={() => zOrder("bringToFront", "Brought to front")}
        >
          ▲▲ Front
        </Button>
        <Button
          size="small"
          disabled={!can1}
          onClick={() => zOrder("bringForward", "Brought forward")}
        >
          ▲ Forward
        </Button>
      </div>
      <div className={styles.row}>
        <Button
          size="small"
          disabled={!can1}
          onClick={() => zOrder("sendToBack", "Sent to back")}
        >
          ▼▼ Back
        </Button>
        <Button
          size="small"
          disabled={!can1}
          onClick={() => zOrder("sendBackward", "Sent backward")}
        >
          ▼ Backward
        </Button>
      </div>
    </div>
  );
};
