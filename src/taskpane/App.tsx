import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  FluentProvider,
  webLightTheme,
  makeStyles,
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Text,
  tokens,
} from "@fluentui/react-components";
import { Toast } from "./components/Toast";
import { AlignPanel } from "./components/AlignPanel";
import { GroupPanel } from "./components/GroupPanel";
import { CommentPanel } from "./components/CommentPanel";
import { SymbolsPanel } from "./components/SymbolsPanel";
import { FootnotePanel } from "./components/FootnotePanel";
import { StatusPanel } from "./components/StatusPanel";
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
  header: {
    padding: "10px 12px 8px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    flexShrink: 0,
  },
  headerTitle: {
    fontWeight: "700",
    fontSize: "14px",
  },
  headerSub: {
    fontSize: "11px",
    opacity: 0.85,
  },
  toastArea: {
    padding: "0 10px",
    flexShrink: 0,
  },
  scrollArea: {
    flex: 1,
    overflowY: "auto",
    padding: "4px 8px 16px",
  },
  accordionPanel: {
    padding: "6px 4px 10px",
  },
});

export const App: React.FC = () => {
  const styles = useStyles();
  const [selectionCount, setSelectionCount] = useState(0);
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

  // Listen for selection changes via Office.js event
  useEffect(() => {
    let handlerId: string | null = null;

    const refresh = async () => {
      try {
        const info = await getSelectionInfo();
        setSelectionCount(info.count);
      } catch {
        setSelectionCount(0);
      }
    };

    // Initial read
    refresh();

    // Register selection change handler
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
      // Fall back to polling if addHandlerAsync not available
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

  const selLabel =
    selectionCount === 0
      ? "Nothing selected"
      : `${selectionCount} shape${selectionCount === 1 ? "" : "s"} selected`;

  return (
    <FluentProvider theme={webLightTheme}>
      <div className={styles.root}>
        {/* Header */}
        <div className={styles.header}>
          <div className={styles.headerTitle}>Slide Toolbox</div>
          <div className={styles.headerSub}>{selLabel}</div>
        </div>

        {/* Toast area */}
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

        {/* Scrollable accordion */}
        <div className={styles.scrollArea}>
          <Accordion multiple collapsible defaultOpenItems={["align"]}>
            <AccordionItem value="align">
              <AccordionHeader>Alignment &amp; Spacing</AccordionHeader>
              <AccordionPanel className={styles.accordionPanel}>
                <AlignPanel
                  selectionCount={selectionCount}
                  showToast={showToast}
                />
              </AccordionPanel>
            </AccordionItem>

            <AccordionItem value="group">
              <AccordionHeader>Grouping &amp; Order</AccordionHeader>
              <AccordionPanel className={styles.accordionPanel}>
                <GroupPanel
                  selectionCount={selectionCount}
                  showToast={showToast}
                />
              </AccordionPanel>
            </AccordionItem>

            <AccordionItem value="comment">
              <AccordionHeader>Sticky Comment</AccordionHeader>
              <AccordionPanel className={styles.accordionPanel}>
                <CommentPanel showToast={showToast} />
              </AccordionPanel>
            </AccordionItem>

            <AccordionItem value="symbols">
              <AccordionHeader>Symbols</AccordionHeader>
              <AccordionPanel className={styles.accordionPanel}>
                <SymbolsPanel showToast={showToast} />
              </AccordionPanel>
            </AccordionItem>

            <AccordionItem value="footnote">
              <AccordionHeader>Footnote / Source</AccordionHeader>
              <AccordionPanel className={styles.accordionPanel}>
                <FootnotePanel showToast={showToast} />
              </AccordionPanel>
            </AccordionItem>

            <AccordionItem value="status">
              <AccordionHeader>Status Label</AccordionHeader>
              <AccordionPanel className={styles.accordionPanel}>
                <StatusPanel showToast={showToast} />
              </AccordionPanel>
            </AccordionItem>
          </Accordion>
        </div>
      </div>
    </FluentProvider>
  );
};
