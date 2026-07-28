/**
 * fontSizes.ts — Font-size cycle helpers.
 *
 * Pure module, no Office.js dependencies. Mirrors the font-size cycle that
 * PowerPoint's "Increase / Decrease Font Size" ribbon buttons step through,
 * so a "notch" in the toolbox UI behaves the way the user expects from the
 * native ribbon.
 */

export const PPT_FONT_SIZE_CYCLE: readonly number[] = [
  8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20,
  24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96,
];

/**
 * Snap `current` to the nearest value in PPT_FONT_SIZE_CYCLE, step `delta`
 * notches, and clamp at the cycle ends.
 *
 * Examples (delta = -2):  16 → 12,  18 → 14,  10 → 8 (clamped).
 * Examples (delta = +2):  12 → 16,  88 → 96 (clamped).
 *
 * Non-cycle inputs (e.g. 13pt) snap to the nearest cycle value first.
 */
export function stepFontSize(current: number, delta: number): number {
  let nearestIdx = 0;
  let nearestDist = Math.abs(PPT_FONT_SIZE_CYCLE[0] - current);
  for (let i = 1; i < PPT_FONT_SIZE_CYCLE.length; i++) {
    const d = Math.abs(PPT_FONT_SIZE_CYCLE[i] - current);
    if (d < nearestDist) {
      nearestDist = d;
      nearestIdx = i;
    }
  }
  const targetIdx = Math.max(
    0,
    Math.min(PPT_FONT_SIZE_CYCLE.length - 1, nearestIdx + delta)
  );
  return PPT_FONT_SIZE_CYCLE[targetIdx];
}
