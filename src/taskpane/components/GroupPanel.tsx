import React from "react";
import { makeStyles, Button, Divider } from "@fluentui/react-components";
import { groupShapes, ungroupShapes, setZOrder, ZOrderType } from "../../lib/ppt";

interface GroupPanelProps {
  selectionCount: number;
  showToast: (msg: string, type: "success" | "error") => void;
}

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "4px" },
  row:  { display: "flex", flexWrap: "wrap", gap: "4px" },
  divider: { margin: "2px 0" },
});

export const GroupPanel: React.FC<GroupPanelProps> = ({ selectionCount, showToast }) => {
  const styles = useStyles();
  const can1 = selectionCount >= 1;
  const can2 = selectionCount >= 2;

  const run = async (fn: () => Promise<void>, msg: string) => {
    try { await fn(); showToast(msg, "success"); }
    catch (e: unknown) { showToast(e instanceof Error ? e.message : String(e), "error"); }
  };

  const z = (type: ZOrderType, label: string) => run(() => setZOrder(type), label);

  return (
    <div className={styles.root}>
      <div className={styles.row}>
        <Button size="small" disabled={!can2} onClick={() => run(groupShapes, "Grouped")}>
          Group
        </Button>
        <Button size="small" disabled={!can1} onClick={() => run(ungroupShapes, "Ungrouped")}>
          Ungroup
        </Button>
      </div>

      <Divider className={styles.divider} />

      <div className={styles.row}>
        <Button size="small" disabled={!can1} onClick={() => z("bringToFront", "Brought to front")}>▲▲ Front</Button>
        <Button size="small" disabled={!can1} onClick={() => z("bringForward",  "Brought forward")} >▲ Forward</Button>
        <Button size="small" disabled={!can1} onClick={() => z("sendToBack",    "Sent to back")}    >▼▼ Back</Button>
        <Button size="small" disabled={!can1} onClick={() => z("sendBackward",  "Sent backward")}  >▼ Backward</Button>
      </div>
    </div>
  );
};
