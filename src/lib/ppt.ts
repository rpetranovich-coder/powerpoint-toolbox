/**
 * ppt.ts — PowerPoint Office.js operation wrappers.
 *
 * All functions use PowerPoint.run() and are async.
 * Errors bubble up as plain Error objects with user-readable messages.
 *
 * Slide dimensions assume widescreen 16:9 default (960 × 540 pt).
 * If your presentation uses 4:3 (720 × 540 pt), update the constants below.
 */

import {
  BoundingBox,
  PositionUpdate,
  computeAlignLeft,
  computeAlignRight,
  computeAlignTop,
  computeAlignBottom,
  computeAlignCenterH,
  computeAlignCenterV,
  computeDistributeH,
  computeDistributeV,
  computeMatchWidth,
  computeMatchHeight,
  computeEqualizeSize,
  computeNudge,
} from "./geometry";

// ─── Constants ────────────────────────────────────────────────────────────────

export const SLIDE_WIDTH_PT = 960;
export const SLIDE_HEIGHT_PT = 540;

// ─── Types ────────────────────────────────────────────────────────────────────

export interface SelectionInfo {
  count: number;
}

// ─── Internal helpers ─────────────────────────────────────────────────────────

/** Load left/top/width/height/name/id on a shape collection and return items. */
async function loadShapes(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.ShapeCollection | PowerPoint.ShapeScopedCollection
): Promise<PowerPoint.Shape[]> {
  shapes.load("items/left,items/top,items/width,items/height,items/name,items/id,items/type");
  await context.sync();
  return shapes.items;
}

/**
 * Get currently selected shapes.  Returns [] if nothing is selected.
 * NOTE: getSelectedShapes() requires PowerPointApi 1.5+ (Microsoft 365).
 */
async function getSelectedShapesItems(
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Shape[]> {
  try {
    return await loadShapes(context, context.presentation.getSelectedShapes());
  } catch {
    return [];
  }
}

/** Get the first selected (or active) slide. */
async function getActiveSlide(
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Slide> {
  const slides = context.presentation.getSelectedSlides();
  slides.load("items");
  await context.sync();
  if (slides.items.length === 0) throw new Error("No slide is active.");
  return slides.items[0];
}

/** Convert an array of shapes into BoundingBox[]. Must be loaded first. */
function shapesToBoxes(shapes: PowerPoint.Shape[]): BoundingBox[] {
  return shapes.map((s) => ({
    left: s.left,
    top: s.top,
    width: s.width,
    height: s.height,
  }));
}

/** Apply PositionUpdates back to an array of shapes. */
function applyUpdates(
  shapes: PowerPoint.Shape[],
  updates: PositionUpdate[]
): void {
  for (const u of updates) {
    const s = shapes[u.index];
    if (u.left !== undefined) s.left = u.left;
    if (u.top !== undefined) s.top = u.top;
    if (u.width !== undefined) s.width = u.width;
    if (u.height !== undefined) s.height = u.height;
  }
}

// ─── Selection info ───────────────────────────────────────────────────────────

export async function getSelectionInfo(): Promise<SelectionInfo> {
  try {
    return await PowerPoint.run(async (context) => {
      const items = await getSelectedShapesItems(context);
      return { count: items.length };
    });
  } catch {
    return { count: 0 };
  }
}

// ─── Alignment ────────────────────────────────────────────────────────────────

export type AlignDirection =
  | "left"
  | "right"
  | "top"
  | "bottom"
  | "centerH"
  | "centerV";

export async function alignShapes(direction: AlignDirection): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapesItems(context);
    if (shapes.length < 2) throw new Error("Select 2 or more shapes to align.");
    const boxes = shapesToBoxes(shapes);
    let updates: PositionUpdate[];
    switch (direction) {
      case "left":     updates = computeAlignLeft(boxes);    break;
      case "right":    updates = computeAlignRight(boxes);   break;
      case "top":      updates = computeAlignTop(boxes);     break;
      case "bottom":   updates = computeAlignBottom(boxes);  break;
      case "centerH":  updates = computeAlignCenterH(boxes); break;
      case "centerV":  updates = computeAlignCenterV(boxes); break;
    }
    applyUpdates(shapes, updates);
    await context.sync();
  });
}

// ─── Distribution ─────────────────────────────────────────────────────────────

export async function distributeShapes(
  direction: "horizontal" | "vertical"
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapesItems(context);
    if (shapes.length < 3)
      throw new Error("Select 3 or more shapes to distribute.");
    const boxes = shapesToBoxes(shapes);
    const updates =
      direction === "horizontal"
        ? computeDistributeH(boxes)
        : computeDistributeV(boxes);
    applyUpdates(shapes, updates);
    await context.sync();
  });
}

// ─── Size matching ────────────────────────────────────────────────────────────

export async function matchSize(
  type: "width" | "height" | "both"
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapesItems(context);
    if (shapes.length < 2)
      throw new Error("Select 2 or more shapes to match size.");
    const boxes = shapesToBoxes(shapes);
    let updates: PositionUpdate[];
    switch (type) {
      case "width":  updates = computeMatchWidth(boxes);   break;
      case "height": updates = computeMatchHeight(boxes);  break;
      case "both":   updates = computeEqualizeSize(boxes); break;
    }
    applyUpdates(shapes, updates);
    await context.sync();
  });
}

// ─── Nudge ────────────────────────────────────────────────────────────────────

export async function nudgeShapes(
  direction: "left" | "right" | "up" | "down",
  amount: number
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapesItems(context);
    if (shapes.length === 0) throw new Error("Select at least one shape to nudge.");
    const boxes = shapesToBoxes(shapes);
    const updates = computeNudge(boxes, direction, amount);
    applyUpdates(shapes, updates);
    await context.sync();
  });
}

// ─── Grouping / Ordering ──────────────────────────────────────────────────────
/**
 * LIMITATION: group() / ungroup() require PowerPointApi 1.6+.
 * On older builds they throw, and we surface a clear error message.
 */

export async function groupShapes(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items");
    await context.sync();
    if (selected.items.length < 2)
      throw new Error("Select 2 or more shapes to group.");
    try {
      selected.group();
      await context.sync();
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      throw new Error(
        `Group not supported on this Office version. (${msg})`
      );
    }
  });
}

export async function ungroupShapes(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapesItems(context);
    if (shapes.length === 0) throw new Error("Select a group to ungroup.");
    const groups = shapes.filter(
      (s) => s.type === PowerPoint.ShapeType.group
    );
    if (groups.length === 0)
      throw new Error("No groups found in selection.");
    try {
      for (const g of groups) {
        g.group.ungroup();
      }
      await context.sync();
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      throw new Error(
        `Ungroup not supported on this Office version. (${msg})`
      );
    }
  });
}

export type ZOrderType =
  | "bringForward"
  | "bringToFront"
  | "sendBackward"
  | "sendToBack";

export async function setZOrder(type: ZOrderType): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelectedShapesItems(context);
    if (shapes.length === 0) throw new Error("Select at least one shape.");
    const map: Record<ZOrderType, PowerPoint.ShapeZOrder> = {
      bringForward:  PowerPoint.ShapeZOrder.bringForward,
      bringToFront:  PowerPoint.ShapeZOrder.bringToFront,
      sendBackward:  PowerPoint.ShapeZOrder.sendBackward,
      sendToBack:    PowerPoint.ShapeZOrder.sendToBack,
    };
    for (const s of shapes) {
      s.setZOrder(map[type]);
    }
    await context.sync();
  });
}

// ─── Sticky Comment ───────────────────────────────────────────────────────────

export async function insertStickyComment(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);

    // Try to find a reference position from selection
    let refLeft = SLIDE_WIDTH_PT / 2 - 110;
    let refTop  = SLIDE_HEIGHT_PT / 2 - 60;

    try {
      const sel = await getSelectedShapesItems(context);
      if (sel.length > 0) {
        // Place to the right of the rightmost selected shape
        const maxRight = Math.max(...sel.map((s) => s.left + s.width));
        const minTop   = Math.min(...sel.map((s) => s.top));
        refLeft = Math.min(maxRight + 16, SLIDE_WIDTH_PT - 240);
        refTop  = minTop + 16;
      }
    } catch {
      // fall through to slide-center default
    }

    const shape = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.rectangle,
      { left: refLeft, top: refTop, width: 220, height: 120 }
    );
    shape.name = `TBX_COMMENT_${Date.now()}`;

    // Fill + border
    shape.fill.setSolidColor("FFF2CC");
    shape.lineFormat.color = "D6B656";
    shape.lineFormat.weight = 1;
    shape.lineFormat.visible = true;

    // Text
    const tf = shape.textFrame;
    tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
    tf.leftMargin  = 6;
    tf.rightMargin = 6;
    tf.topMargin   = 4;

    const range = tf.textRange;
    range.text = "COMMENT: <type here>";
    range.font.name = "Segoe UI";
    range.font.size = 12;
    range.font.color = "333333";
    range.font.bold  = false;

    await context.sync();
  });
}

// ─── Symbol insertion ─────────────────────────────────────────────────────────
/**
 * Insert an SVG symbol using addSvgImage().
 * REQUIREMENT: Microsoft 365 / Office 2021+ on Windows.
 * Falls back with a clear error on older versions.
 */
export async function insertSymbol(
  svgContent: string,
  sizePt: number
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);

    let left = SLIDE_WIDTH_PT / 2 - sizePt / 2;
    let top  = SLIDE_HEIGHT_PT / 2 - sizePt / 2;

    try {
      const sel = await getSelectedShapesItems(context);
      if (sel.length > 0) {
        const maxRight = Math.max(...sel.map((s) => s.left + s.width));
        const minTop   = Math.min(...sel.map((s) => s.top));
        left = Math.min(maxRight + 8, SLIDE_WIDTH_PT - sizePt - 8);
        top  = minTop;
      }
    } catch {
      // use center default
    }

    try {
      // addSvgImage is available in PowerPointApi 1.6+ but not yet in @types/office-js.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (slide.shapes as any).addSvgImage(svgContent, {
        left,
        top,
        width: sizePt,
        height: sizePt,
      });
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      throw new Error(
        `SVG insertion failed — requires Microsoft 365 / Office 2021+. (${msg})`
      );
    }

    await context.sync();
  });
}

// ─── Footnote / Source ────────────────────────────────────────────────────────

const FOOTNOTE_NAME = "TBX_FOOTNOTE";
const SOURCE_NAME   = "TBX_SOURCE";
const BOTTOM_MARGIN = 14;
const LEFT_MARGIN   = 14;
const NOTE_WIDTH    = Math.round(SLIDE_WIDTH_PT * 0.88);
const NOTE_HEIGHT   = 22;

async function upsertBottomTextBox(
  prefix: string,
  text: string,
  shapeName: string,
  topOffset: number
): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    slide.shapes.load("items/name");
    await context.sync();

    const existing = slide.shapes.items.find((s) => s.name === shapeName);

    if (existing) {
      existing.load("textFrame/textRange/text");
      await context.sync();
      existing.textFrame.textRange.text = `${prefix}${text}`;
      await context.sync();
      return;
    }

    const topPt =
      SLIDE_HEIGHT_PT - NOTE_HEIGHT - BOTTOM_MARGIN - topOffset;

    const shape = slide.shapes.addTextBox(`${prefix}${text}`, {
      left:   LEFT_MARGIN,
      top:    topPt,
      width:  NOTE_WIDTH,
      height: NOTE_HEIGHT,
    });
    shape.name = shapeName;

    const tf = shape.textFrame;
    tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
    tf.leftMargin  = 0;
    tf.rightMargin = 0;
    tf.topMargin   = 0;

    shape.lineFormat.visible = false;

    const range = tf.textRange;
    range.font.name  = "Segoe UI";
    range.font.size  = 8.5;
    range.font.color = "666666";
    range.font.bold  = false;

    await context.sync();
  });
}

export async function insertFootnote(text: string): Promise<void> {
  await upsertBottomTextBox("Note: ", text, FOOTNOTE_NAME, 0);
}

export async function insertSource(text: string): Promise<void> {
  // Source sits one row above footnote
  await upsertBottomTextBox("Source: ", text, SOURCE_NAME, NOTE_HEIGHT + 4);
}

// ─── Status Label ─────────────────────────────────────────────────────────────

const STATUS_NAME  = "TBX_STATUS";
const STATUS_W     = 220;
const STATUS_H     = 22;
const STATUS_RIGHT = 20;
const STATUS_TOP   = 18;

export async function insertStatusLabel(text: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    slide.shapes.load("items/name");
    await context.sync();

    const existing = slide.shapes.items.find((s) => s.name === STATUS_NAME);

    if (existing) {
      existing.load("textFrame/textRange/text");
      await context.sync();
      existing.textFrame.textRange.text = text;
      await context.sync();
      return;
    }

    const left = SLIDE_WIDTH_PT - STATUS_W - STATUS_RIGHT;

    const shape = slide.shapes.addTextBox(text, {
      left,
      top:    STATUS_TOP,
      width:  STATUS_W,
      height: STATUS_H,
    });
    shape.name = STATUS_NAME;

    const tf = shape.textFrame;
    tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
    tf.leftMargin  = 0;
    tf.rightMargin = 0;
    tf.topMargin   = 0;

    shape.lineFormat.visible = false;

    const range = tf.textRange;
    range.font.name  = "Segoe UI Semibold";
    range.font.size  = 11;
    range.font.color = "808080";
    range.font.bold  = true;

    tf.textRange.paragraphFormat.horizontalAlignment =
      PowerPoint.ParagraphHorizontalAlignment.right;

    await context.sync();
  });
}
