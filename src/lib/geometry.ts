/**
 * geometry.ts — Pure math helpers for shape alignment and distribution.
 * No Office.js dependencies; fully unit-testable.
 */

export interface BoundingBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface PositionUpdate {
  index: number;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

// ─── Alignment ────────────────────────────────────────────────────────────────

/** Align all shapes to the leftmost left edge. */
export function computeAlignLeft(boxes: BoundingBox[]): PositionUpdate[] {
  const anchor = Math.min(...boxes.map((b) => b.left));
  return boxes.map((_, i) => ({ index: i, left: anchor }));
}

/** Align all right edges to the rightmost right edge. */
export function computeAlignRight(boxes: BoundingBox[]): PositionUpdate[] {
  const anchor = Math.max(...boxes.map((b) => b.left + b.width));
  return boxes.map((b, i) => ({ index: i, left: anchor - b.width }));
}

/** Align all shapes to the topmost top edge. */
export function computeAlignTop(boxes: BoundingBox[]): PositionUpdate[] {
  const anchor = Math.min(...boxes.map((b) => b.top));
  return boxes.map((_, i) => ({ index: i, top: anchor }));
}

/** Align all bottom edges to the bottommost bottom edge. */
export function computeAlignBottom(boxes: BoundingBox[]): PositionUpdate[] {
  const anchor = Math.max(...boxes.map((b) => b.top + b.height));
  return boxes.map((b, i) => ({ index: i, top: anchor - b.height }));
}

/** Center all shapes horizontally within their collective bounding box. */
export function computeAlignCenterH(boxes: BoundingBox[]): PositionUpdate[] {
  const minLeft = Math.min(...boxes.map((b) => b.left));
  const maxRight = Math.max(...boxes.map((b) => b.left + b.width));
  const centerX = (minLeft + maxRight) / 2;
  return boxes.map((b, i) => ({ index: i, left: centerX - b.width / 2 }));
}

/** Center all shapes vertically within their collective bounding box. */
export function computeAlignCenterV(boxes: BoundingBox[]): PositionUpdate[] {
  const minTop = Math.min(...boxes.map((b) => b.top));
  const maxBottom = Math.max(...boxes.map((b) => b.top + b.height));
  const centerY = (minTop + maxBottom) / 2;
  return boxes.map((b, i) => ({ index: i, top: centerY - b.height / 2 }));
}

// ─── Edge-based distribution ──────────────────────────────────────────────────

/**
 * Distribute shapes horizontally with equal gaps between edges.
 * The leftmost shape's left and rightmost shape's right are fixed.
 * Requires n >= 3 shapes; returns unmodified indices for n < 3.
 */
export function computeDistributeH(boxes: BoundingBox[]): PositionUpdate[] {
  if (boxes.length < 3) return boxes.map((_, i) => ({ index: i }));

  const sorted = boxes
    .map((b, i) => ({ ...b, origIndex: i }))
    .sort((a, b) => a.left - b.left);

  const totalSpan =
    sorted[sorted.length - 1].left +
    sorted[sorted.length - 1].width -
    sorted[0].left;
  const totalWidths = sorted.reduce((s, b) => s + b.width, 0);
  const gap = (totalSpan - totalWidths) / (sorted.length - 1);

  const updates: PositionUpdate[] = [];
  let cursor = sorted[0].left;
  for (const b of sorted) {
    updates.push({ index: b.origIndex, left: cursor });
    cursor += b.width + gap;
  }
  return updates;
}

/**
 * Distribute shapes vertically with equal gaps between edges.
 * The topmost shape's top and bottommost shape's bottom are fixed.
 * Requires n >= 3.
 */
export function computeDistributeV(boxes: BoundingBox[]): PositionUpdate[] {
  if (boxes.length < 3) return boxes.map((_, i) => ({ index: i }));

  const sorted = boxes
    .map((b, i) => ({ ...b, origIndex: i }))
    .sort((a, b) => a.top - b.top);

  const totalSpan =
    sorted[sorted.length - 1].top +
    sorted[sorted.length - 1].height -
    sorted[0].top;
  const totalHeights = sorted.reduce((s, b) => s + b.height, 0);
  const gap = (totalSpan - totalHeights) / (sorted.length - 1);

  const updates: PositionUpdate[] = [];
  let cursor = sorted[0].top;
  for (const b of sorted) {
    updates.push({ index: b.origIndex, top: cursor });
    cursor += b.height + gap;
  }
  return updates;
}

// ─── Size matching ────────────────────────────────────────────────────────────

/** Set all shapes' widths to the anchor shape's width. Anchor = index 0. */
export function computeMatchWidth(
  boxes: BoundingBox[],
  anchorIndex = 0
): PositionUpdate[] {
  const w = boxes[anchorIndex].width;
  return boxes.map((_, i) => ({ index: i, width: w }));
}

/** Set all shapes' heights to the anchor shape's height. Anchor = index 0. */
export function computeMatchHeight(
  boxes: BoundingBox[],
  anchorIndex = 0
): PositionUpdate[] {
  const h = boxes[anchorIndex].height;
  return boxes.map((_, i) => ({ index: i, height: h }));
}

/** Set all shapes to the anchor shape's width AND height. Anchor = index 0. */
export function computeEqualizeSize(
  boxes: BoundingBox[],
  anchorIndex = 0
): PositionUpdate[] {
  const { width: w, height: h } = boxes[anchorIndex];
  return boxes.map((_, i) => ({ index: i, width: w, height: h }));
}

// ─── Nudge ────────────────────────────────────────────────────────────────────

export function computeNudge(
  boxes: BoundingBox[],
  direction: "left" | "right" | "up" | "down",
  amount: number
): PositionUpdate[] {
  return boxes.map((b, i) => {
    switch (direction) {
      case "left":
        return { index: i, left: b.left - amount };
      case "right":
        return { index: i, left: b.left + amount };
      case "up":
        return { index: i, top: b.top - amount };
      case "down":
        return { index: i, top: b.top + amount };
    }
  });
}
