# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev-server         # Start HTTPS dev server on https://localhost:3000
npm run start              # Sideload + open in PowerPoint (web/desktop auto-detect)
npm run start:desktop      # Force sideload into desktop PowerPoint
npm run stop               # Stop sideloaded add-in
npm run build              # Production build → dist/
npm run build:dev          # Dev build → dist/
npm run validate           # Validate manifest.xml schema
npm run lint               # ESLint .ts / .tsx
npm run lint:fix           # Auto-fix lint issues
```

## Architecture

**PowerPoint Office Add-in** — TypeScript + React 18 + Fluent UI v9 + Webpack.
Entry: `src/taskpane/index.tsx` → mounts React inside `Office.onReady()`.

### Key directories

| Path | Purpose |
|---|---|
| `src/lib/geometry.ts` | **Pure math** — alignment, edge-based distribution, nudge. No Office deps; fully unit-testable. |
| `src/lib/ppt.ts` | **All Office.js operations.** Imports geometry helpers. Exposes async functions called by React components. |
| `src/lib/symbols.ts` | Inline SVG strings for Harvey balls, stoplights, arrows. |
| `src/taskpane/App.tsx` | Root React component. Manages selection count (via `DocumentSelectionChanged` event) and toast state. |
| `src/taskpane/components/` | One panel component per feature section (Align, Group, Comment, Symbols, Footnote, Status). |
| `manifest.xml` | Add-in identity, ribbon button ("Open Toolbox" on Home tab), resource URLs pointing to `localhost:3000`. |

### Office.js API notes

- `context.presentation.getSelectedShapes()` → `ShapeScopedCollection` (PowerPointApi 1.5+)
- Z-order enum: `PowerPoint.ShapeZOrder` (not `ShapeZOrderType`)
- `addSvgImage` exists at runtime (PowerPointApi 1.6+) but is absent from `@types/office-js`; called via `(slide.shapes as any).addSvgImage(...)`
- Paragraph alignment property: `paragraphFormat.horizontalAlignment` (not `.alignment`)
- Slide default: 960 × 540 pt (widescreen 16:9) — update `SLIDE_WIDTH_PT`/`SLIDE_HEIGHT_PT` in `ppt.ts` for 4:3

### Shape naming convention

| Shape | Name pattern |
|---|---|
| Sticky comment | `TBX_COMMENT_<timestamp>` |
| Footnote | `TBX_FOOTNOTE` |
| Source | `TBX_SOURCE` |
| Status label | `TBX_STATUS` |

Footnote/Source/Status are upserted: if a shape with that name exists on the current slide, its text is updated instead of creating a duplicate.
