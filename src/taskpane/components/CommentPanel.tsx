import React from "react";
import { makeStyles, Button, Text, tokens } from "@fluentui/react-components";
import { Note20Regular } from "@fluentui/react-icons";
import { insertStickyComment } from "../../lib/ppt";

interface CommentPanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "6px" },
  hint: { color: tokens.colorNeutralForeground3, fontSize: "11px" },
});

export const CommentPanel: React.FC<CommentPanelProps> = ({ showToast }) => {
  const styles = useStyles();

  const handleInsert = async () => {
    try {
      await insertStickyComment();
      showToast("Sticky comment inserted", "success");
    } catch (e: unknown) {
      showToast(e instanceof Error ? e.message : String(e), "error");
    }
  };

  return (
    <div className={styles.root}>
      <Text className={styles.hint}>
        Inserts a yellow sticky-note shape. Named{" "}
        <code>TBX_COMMENT_*</code> — editable after insertion.
      </Text>
      <Button size="small" icon={<Note20Regular />} onClick={handleInsert}>
        Add Sticky Comment
      </Button>
    </div>
  );
};
