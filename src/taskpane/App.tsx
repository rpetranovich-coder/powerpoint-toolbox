import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  FluentProvider,
  webDarkTheme,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { Toast } from "./components/Toast";
import { AlignPanel } from "./components/AlignPanel";
import { GroupPanel } from "./components/GroupPanel";
import { CommentPanel } from "./components/CommentPanel";
import { SymbolsPanel } from "./components/SymbolsPanel";
import { FootnotePanel } from "./components/FootnotePanel";
import { StatusPanel } from "./components/StatusPanel";
import { FontPanel } from "./components/FontPanel";
import { TablePanel } from "./components/TablePanel";
import { GuidesPanel } from "./components/GuidesPanel";
import { AiPanel } from "./components/AiPanel";
import { getSelectionInfo } from "../lib/ppt";

interface ToastState {
  id: number;
  message: string;
  type: "success" | "error";
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
  },

  // ── Toast ───────────────────────────────────────────────────────────────
  toastArea: {
    flexShrink: 0,
    padding: "0 10px",
  },

  // ── Scroll container ────────────────────────────────────────────────────
  scrollArea: {
    flex: 1,
    overflowY: "auto",
    // Custom scrollbar — matches Office task panes
    scrollbarWidth: "thin",
    scrollbarColor: `${tokens.colorNeutralStroke1} transparent`,
  },

  // ── Section divider ──────────────────────────────────────────────────────
  sectionDivider: {
    border: "none",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    margin: "0",
  },

  // ── Panel body ───────────────────────────────────────────────────────────
  panelBody: {
    padding: "6px 8px",
  },
});

export const App: React.FC = () => {
  const styles = useStyles();

  const [selectionCount, setSelectionCount] = useState(0);
  const [selectedCommentName, setSelectedCommentName] = useState<string | null>(null);
  const [selectedStoplightName, setSelectedStoplightName] = useState<string | null>(null);
  const [toast, setToast] = useState<ToastState | null>(null);
  const toastIdRef = useRef(0);

  const showToast = useCallback(
    (message: string, type: "success" | "error") => {
      toastIdRef.current += 1;
      setToast({ id: toastIdRef.current, message, type });
    },
    []
  );

  const dismissToast = useCallback(() => setToast(null), []);

  useEffect(() => {
    let handlerId: string | null = null;

    const refresh = async () => {
      try {
        const info = await getSelectionInfo();
        setSelectionCount(info.count);
        setSelectedCommentName(info.selectedCommentName);
        setSelectedStoplightName(info.selectedStoplightName);
      } catch {
        setSelectionCount(0);
        setSelectedCommentName(null);
        setSelectedStoplightName(null);
      }
    };

    refresh();

    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        refresh,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            handlerId = "registered";
          }
        }
      );
    } catch {
      const interval = setInterval(refresh, 1500);
      return () => clearInterval(interval);
    }

    return () => {
      if (handlerId) {
        try {
          Office.context.document.removeHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            { handler: refresh },
            () => { /* no-op */ }
          );
        } catch {
          // ignore
        }
      }
    };
  }, []);

  return (
    <FluentProvider theme={webDarkTheme}>
      <div className={styles.root}>

        {/* ── Toast ── */}
        <div className={styles.toastArea}>
          {toast && (
            <Toast
              key={toast.id}
              message={toast.message}
              type={toast.type}
              onDismiss={dismissToast}
            />
          )}
        </div>

        {/* ── Flat sections ── */}
        <div className={styles.scrollArea}>

          <div className={styles.panelBody}>
            <AiPanel showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <AlignPanel selectionCount={selectionCount} showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <GroupPanel selectionCount={selectionCount} showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <CommentPanel selectedCommentName={selectedCommentName} selectedStoplightName={selectedStoplightName} showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <SymbolsPanel showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <TablePanel showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <FontPanel showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <GuidesPanel showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <FootnotePanel showToast={showToast} />
          </div>

          <hr className={styles.sectionDivider} />
          <div className={styles.panelBody}>
            <StatusPanel showToast={showToast} />
          </div>

        </div>
      </div>
    </FluentProvider>
  );
};
