import React, { useState } from "react";
import {
  makeStyles,
  Button,
  ToggleButton,
  Divider,
  Tooltip,
  tokens,
} from "@fluentui/react-components";
import { insertSymbol } from "../../lib/ppt";
import { ALL_SYMBOL_GROUPS, Symbol } from "../../lib/symbols";

interface SymbolsPanelProps {
  showToast: (msg: string, type: "success" | "error") => void;
}

const SIZE_OPTIONS: { label: string; pt: number }[] = [
  { label: "S", pt: 18 },
  { label: "M", pt: 28 },
  { label: "L", pt: 40 },
];

const useStyles = makeStyles({
  root:      { display: "flex", flexDirection: "column", gap: "4px" },
  sizeRow:   { display: "flex", gap: "4px", alignItems: "center" },
  symbolRow: { display: "flex", flexWrap: "wrap", gap: "3px" },
  symbolBtn: { minWidth: "28px", padding: "2px 3px" },
  divider:   { margin: "2px 0" },
  sizeLabel: { fontSize: "10px", color: tokens.colorNeutralForeground3 },
});

const SvgIcon: React.FC<{ svg: string; size?: number }> = ({ svg, size = 18 }) => (
  <span
    style={{ display: "inline-flex", width: size, height: size }}
    dangerouslySetInnerHTML={{ __html: svg }}
  />
);

export const SymbolsPanel: React.FC<SymbolsPanelProps> = ({ showToast }) => {
  const styles = useStyles();
  const [sizePt, setSizePt] = useState(28);

  const handleInsert = async (symbol: Symbol) => {
    try { await insertSymbol(symbol.svg, sizePt); showToast(`${symbol.label} inserted`, "success"); }
    catch (e: unknown) { showToast(e instanceof Error ? e.message : String(e), "error"); }
  };

  return (
    <div className={styles.root}>
      <div className={styles.sizeRow}>
        <span className={styles.sizeLabel}>Size:</span>
        {SIZE_OPTIONS.map((opt) => (
          <ToggleButton key={opt.pt} size="small" checked={sizePt === opt.pt} onClick={() => setSizePt(opt.pt)}>
            {opt.label}
          </ToggleButton>
        ))}
      </div>

      {ALL_SYMBOL_GROUPS.map((group, gi) => (
        <React.Fragment key={group.groupId}>
          {gi > 0 && <Divider className={styles.divider} />}
          <div className={styles.symbolRow}>
            {group.symbols.map((sym) => (
              <Tooltip key={sym.id} content={sym.label} relationship="label">
                <Button className={styles.symbolBtn} size="small"
                  icon={<SvgIcon svg={sym.svg} size={16} />}
                  onClick={() => handleInsert(sym)} />
              </Tooltip>
            ))}
          </div>
        </React.Fragment>
      ))}
    </div>
  );
};
