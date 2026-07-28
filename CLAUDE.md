# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run deploy             # ⭐ Build + publish to gh-pages (THIS is how changes go live in PowerPoint)
npm run build              # Production build → dist/
npm run build:dev          # Dev build → dist/
npm run dev-server         # HTTPS dev server on https://localhost:3000 (NOT used by this project's manifest — see below)
npm run start              # Sideload + open in PowerPoint (web/desktop auto-detect)
npm run start:desktop      # Force sideload into desktop PowerPoint
npm run stop               # Stop sideloaded add-in
npm run validate           # Validate manifest.xml schema
npm run lint               # ESLint .ts / .tsx
npm run lint:fix           # Auto-fix lint issues
```

## How updates get into PowerPoint (READ THIS BEFORE TELLING THE USER HOW TO TEST)

**This project's `manifest.xml` points all resource URLs at `https://rpetranovich-coder.github.io/powerpoint-toolbox/` (GitHub Pages), NOT at `localhost:3000`.** The PowerPoint add-in always loads the bundle from GitHub Pages, regardless of whether a dev server is running. `localhost:3000` is used **only** for local dev via `npm run dev-server`; it is not what the sideloaded add-in normally loads.

The normal workflow to see code changes in PowerPoint:
1. Edit files under `src/**`.
2. `npm run deploy` — runs `npm run build` (webpack → `dist/`) then `gh-pages -d dist` (pushes `dist/` to the `gh-pages` branch). GitHub Pages serves it within ~1 min.
3. In PowerPoint: close the toolbox task pane and reopen it (Home → Open Toolbox). The new bundle loads. No re-sideload is needed for code changes because the manifest references remote URLs.

**Do NOT default to suggesting `npm run dev-server` + `npm run start:desktop`** — that workflow requires a localhost-pointing manifest, which this project does not have. If the user ever wants live-reload dev iteration, they would need a separate localhost-pointing dev manifest. Don't propose this unless asked.

### When a re-sideload IS required

- Only changes to `manifest.xml` itself (ribbon buttons, permissions, IDs, URLs, task-pane names) require re-uploading the manifest in PowerPoint. Code/asset changes served from GitHub Pages do not.

### Caching

- PowerPoint's webview caches `taskpane.html`/JS. If a change doesn't appear after reopening, clear the Office cache at `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` or fully restart PowerPoint.

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
| `manifest.xml` | Add-in identity, ribbon button ("Open Toolbox" on Home tab), resource URLs pointing to GitHub Pages (`https://rpetranovich-coder.github.io/powerpoint-toolbox/`). |

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
| Bullets box | `TBX_BULLETS_<timestamp>` |
| Footnote | `TBX_FOOTNOTE` |
| Source | `TBX_SOURCE` |
| Status label | `TBX_STATUS` |

Footnote/Source/Status are upserted: if a shape with that name exists on the current slide, its text is updated instead of creating a duplicate. Comment and Bullets shapes use unique timestamped names so multiple can coexist per slide.
