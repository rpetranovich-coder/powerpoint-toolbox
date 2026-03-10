/**
 * symbols.ts — Inline SVG definitions for the symbol palette.
 * All symbols use viewBox="0 0 24 24" so they scale cleanly.
 */

export interface Symbol {
  id: string;
  label: string;
  svg: string;
}

// ─── Harvey Balls ─────────────────────────────────────────────────────────────
// Fill clockwise from 12-o'clock.  Outline circle always drawn on top.

const HARVEY_FILL = "#4a4a4a";
const HARVEY_STROKE = "#4a4a4a";
const HARVEY_BG = "#ffffff";

// cx=12, cy=12, r=10  → top=(12,2)  right=(22,12)  bottom=(12,22)  left=(2,12)
const HARVEY_CIRCLE_OUTLINE = `<circle cx="12" cy="12" r="10" fill="none" stroke="${HARVEY_STROKE}" stroke-width="1.5"/>`;

const HARVEY_BALLS: Symbol[] = [
  {
    id: "harvey_0",
    label: "0%",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" fill="${HARVEY_BG}" stroke="${HARVEY_STROKE}" stroke-width="1.5"/>
    </svg>`,
  },
  {
    id: "harvey_25",
    label: "25%",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" fill="${HARVEY_BG}"/>
      <!-- top-right quadrant: center→top arc90°CW→right→center -->
      <path d="M12,12 L12,2 A10,10 0 0,1 22,12 Z" fill="${HARVEY_FILL}"/>
      ${HARVEY_CIRCLE_OUTLINE}
    </svg>`,
  },
  {
    id: "harvey_50",
    label: "50%",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" fill="${HARVEY_BG}"/>
      <!-- right half: center→top arc180°CW→bottom→center -->
      <path d="M12,12 L12,2 A10,10 0 1,1 12,22 Z" fill="${HARVEY_FILL}"/>
      ${HARVEY_CIRCLE_OUTLINE}
    </svg>`,
  },
  {
    id: "harvey_75",
    label: "75%",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" fill="${HARVEY_BG}"/>
      <!-- 3/4: center→top arc270°CW→left→center -->
      <path d="M12,12 L12,2 A10,10 0 1,1 2,12 Z" fill="${HARVEY_FILL}"/>
      ${HARVEY_CIRCLE_OUTLINE}
    </svg>`,
  },
  {
    id: "harvey_100",
    label: "100%",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" fill="${HARVEY_FILL}" stroke="${HARVEY_STROKE}" stroke-width="1.5"/>
    </svg>`,
  },
];

// ─── Stoplights ───────────────────────────────────────────────────────────────

const STOPLIGHTS: Symbol[] = [
  {
    id: "stoplight_red",
    label: "Red",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" fill="#CC2222" stroke="#881111" stroke-width="1"/>
    </svg>`,
  },
  {
    id: "stoplight_yellow",
    label: "Yellow",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" fill="#FFCC00" stroke="#AA8800" stroke-width="1"/>
    </svg>`,
  },
  {
    id: "stoplight_green",
    label: "Green",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" fill="#22AA22" stroke="#116611" stroke-width="1"/>
    </svg>`,
  },
];

// ─── Arrows ───────────────────────────────────────────────────────────────────

const ARROW_FILL = "#333333";

const ARROWS: Symbol[] = [
  {
    id: "arrow_right",
    label: "→",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <path d="M4,10 H16 L12,5 L20,12 L12,19 L16,14 H4 Z" fill="${ARROW_FILL}"/>
    </svg>`,
  },
  {
    id: "arrow_left",
    label: "←",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <path d="M20,10 H8 L12,5 L4,12 L12,19 L8,14 H20 Z" fill="${ARROW_FILL}"/>
    </svg>`,
  },
  {
    id: "arrow_up",
    label: "↑",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <path d="M14,20 V8 L19,12 L12,4 L5,12 L10,8 V20 Z" fill="${ARROW_FILL}"/>
    </svg>`,
  },
  {
    id: "arrow_down",
    label: "↓",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <path d="M10,4 V16 L5,12 L12,20 L19,12 L14,16 V4 Z" fill="${ARROW_FILL}"/>
    </svg>`,
  },
  {
    id: "arrow_curved_right",
    label: "↪",
    svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
      <path d="M4,16 C4,10 9,6 15,7 L15,4 L21,9 L15,14 L15,11 C10,10 8,13 8,16 Z" fill="${ARROW_FILL}"/>
    </svg>`,
  },
];

export { HARVEY_BALLS, STOPLIGHTS, ARROWS };

export const ALL_SYMBOL_GROUPS = [
  { groupId: "harvey", label: "Harvey Balls", symbols: HARVEY_BALLS },
  { groupId: "stoplight", label: "Stoplights", symbols: STOPLIGHTS },
  { groupId: "arrows", label: "Arrows", symbols: ARROWS },
];
