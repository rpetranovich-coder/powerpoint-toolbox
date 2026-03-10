import React, { useEffect } from "react";
import { makeStyles, tokens, Text } from "@fluentui/react-components";
import {
  CheckmarkCircle20Filled,
  ErrorCircle20Filled,
} from "@fluentui/react-icons";

interface ToastProps {
  message: string;
  type: "success" | "error";
  onDismiss: () => void;
}

const useStyles = makeStyles({
  toast: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "8px 10px",
    borderRadius: "4px",
    margin: "6px 0",
  },
  success: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground2,
  },
  error: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground2,
  },
});

export const Toast: React.FC<ToastProps> = ({ message, type, onDismiss }) => {
  const styles = useStyles();

  useEffect(() => {
    const t = setTimeout(onDismiss, 3500);
    return () => clearTimeout(t);
  }, [message, onDismiss]);

  return (
    <div
      className={`${styles.toast} ${type === "success" ? styles.success : styles.error}`}
    >
      {type === "success" ? (
        <CheckmarkCircle20Filled />
      ) : (
        <ErrorCircle20Filled />
      )}
      <Text size={200}>{message}</Text>
    </div>
  );
};
