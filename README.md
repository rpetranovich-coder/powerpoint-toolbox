# Slide Toolbox — PowerPoint Add-in

A consulting-grade task-pane add-in for PowerPoint on Windows desktop.
Built with TypeScript + React + Fluent UI v9 + Office.js.

---

## Features

| Section | What it does |
|---|---|
| **Alignment** | Align edges, distribute (edge-based, even gaps), match/equalize size, nudge 1pt/5pt |
| **Grouping & Order** | Group, Ungroup, Bring to Front/Forward, Send to Back/Backward |
| **Sticky Comment** | Insert a yellow sticky-note shape (`TBX_COMMENT_*`) editable after insertion |
| **Symbols** | Harvey balls (0/25/50/75/100%), Stoplights (R/Y/G), Arrows — inserted as SVG at S/M/L sizes |
| **Footnote / Source** | Insert/update `Note:` and `Source:` text boxes at bottom-left; no duplicates |
| **Status Label** | Insert/update `/PRELIMINARY`, `/DRAFT`, `/CONFIDENTIAL`, `/INTERNAL` at top-right |

---

## Prerequisites

- **Node.js** ≥ 18
- **npm** ≥ 9
- **PowerPoint for Windows** (Microsoft 365 / Office 2021 recommended)
- A modern browser for the dev-server HTTPS certificate (Edge or Chrome)

---

## Install & run

```bash
cd powerpoint-toolbox
npm install
npm run dev-server        # starts HTTPS on https://localhost:3000
```

In a **second terminal** (to sideload):

```bash
npm run start             # sideloads into PowerPoint (web)
# OR
npm run start:desktop     # sideloads into PowerPoint desktop
```

The first time you run the dev-server, Office Add-in Dev Certs will prompt you to trust a self-signed certificate.
Accept it (or run `npx office-addin-dev-certs install` manually first).

---

## Sideload manually (Windows desktop PowerPoint)

1. Open PowerPoint → **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
2. Add the catalog URL **`https://localhost:3000`** and tick *Show in Menu*.
   *(Or use the simpler route below.)*
3. Alternatively: **Insert → Get Add-ins → Upload My Add-in** → browse to `manifest.xml`.
4. The **"Open Toolbox"** button appears in the **Home** ribbon tab.

---

## Debugging

- Open DevTools in the task pane: right-click inside the pane → **Inspect** (requires `--enable-runtime-exceptions` flag or use Edge WebView2 debug port).
- Or attach VS Code debugger: install the *Office Add-ins Debugger* extension and press F5.
- Console logs and errors appear in the task pane DevTools console.

---

## Build for production

```bash
npm run build          # outputs to /dist
```

Replace all `https://localhost:3000` references in `manifest.xml` with your hosted URL before deploying.

---

## Known Office.js limitations & workarounds

### Selection API (`getSelectedShapes`)
- Requires **PowerPointApi 1.5+** (Microsoft 365, builds ≥ 2108).
- On older builds, the selection count will read as 0 and alignment buttons will stay disabled.
- Workaround: upgrade to Microsoft 365.

### Group / Ungroup (`ShapeCollection.group()`)
- Requires **PowerPointApi 1.6+**.
- On older builds the button shows a clear error toast instead of crashing.

### SVG symbol insertion (`addSvgImage`)
- Requires **PowerPointApi 1.6+** / Microsoft 365.
- On older Office the symbol buttons will toast an error explaining the requirement.

### Slide dimensions
- Default assumes widescreen 16:9 (960 × 540 pt).
- If your presentation uses 4:3 (720 × 540 pt), update `SLIDE_WIDTH_PT` in `src/lib/ppt.ts`.
- A future improvement would be to read dimensions dynamically via `context.presentation` when the API supports it.

### Real PowerPoint comments
- True PowerPoint comments are not accessible via Office.js on desktop as of 2025.
- The "Sticky Comment" feature uses a shape-based approach (a yellow rectangle) as a workaround.

### `shape.group.ungroup()` error handling
- If the selected shape is not a group, a user-friendly error toast is shown.

---

## File structure

```
powerpoint-toolbox/
├── manifest.xml                  # Add-in manifest (IDs, ribbon entry, URLs)
├── webpack.config.js             # Dev server + bundler
├── src/
│   ├── lib/
│   │   ├── geometry.ts           # Pure alignment/distribution math (no Office deps)
│   │   ├── ppt.ts                # All Office.js PowerPoint operations
│   │   └── symbols.ts            # Inline SVG definitions for the symbol palette
│   ├── taskpane/
│   │   ├── index.tsx             # React entry point (Office.onReady → ReactDOM)
│   │   ├── App.tsx               # Root component, selection polling, toast state
│   │   ├── taskpane.html         # HTML shell (loads office.js + webpack bundle)
│   │   ├── taskpane.css          # Global resets
│   │   └── components/
│   │       ├── Toast.tsx
│   │       ├── AlignPanel.tsx
│   │       ├── GroupPanel.tsx
│   │       ├── CommentPanel.tsx
│   │       ├── SymbolsPanel.tsx
│   │       ├── FootnotePanel.tsx
│   │       └── StatusPanel.tsx
│   └── commands/
│       ├── commands.html
│       └── commands.ts           # Placeholder for future ribbon commands
└── assets/                       # Icon PNGs referenced by manifest
```

---

## Next improvements

- Read actual slide dimensions dynamically instead of using constants
- Add recoloring option for inserted symbols (fill override via shape properties)
- Support shape tag reading to list/delete all TBX_* shapes on the current slide
- Add a "Clear all TBX shapes" utility
- Package for AppSource / SharePoint catalog deployment
