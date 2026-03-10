# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev-server      # Start local HTTPS dev server on port 3000
npm run start           # Sideload add-in into PowerPoint (web)
npm run start:desktop   # Sideload add-in into PowerPoint (desktop)
npm run stop            # Stop the sideloaded add-in
npm run build           # Production build to dist/
npm run build:dev       # Development build to dist/
npm run validate        # Validate manifest.xml schema
npm run lint            # ESLint TypeScript files
npm run lint:fix        # Auto-fix lint issues
```

## Architecture

This is a **PowerPoint Office Add-in** using the [Office.js](https://learn.microsoft.com/office/dev/add-ins/) platform (TypeScript + Webpack).

### Key concepts

- **manifest.xml** — declares the add-in identity, permissions, and where to load HTML/JS from. The dev server URL (`https://localhost:3000`) is hardcoded here; update all URLs before production deployment.
- **Task pane** (`src/taskpane/`) — the sidebar UI that opens in PowerPoint. Entry point is `taskpane.ts`, which calls `Office.onReady()` before touching any Office APIs.
- **Commands** (`src/commands/`) — background JS for ribbon button actions that run without opening the task pane.
- **assets/** — icon files referenced in the manifest (required sizes: 16, 32, 64, 80 px PNG).

### Office.js API entry point

All Office API calls must happen inside the `Office.onReady()` callback (or after `await Office.onReady()`). Use `PowerPoint.run(async (context) => { ... })` for the PowerPoint-specific API with auto-commit.

### Dev workflow

1. Run `npm run dev-server` to start the HTTPS server.
2. Open PowerPoint and sideload via `npm run start` (or manually via **Insert > Add-ins > Upload My Add-in** using `manifest.xml`).
3. The task pane loads from `https://localhost:3000/taskpane.html`.

### Build output

Webpack bundles `src/taskpane/taskpane.ts` and `src/commands/commands.ts` into `dist/`, along with the HTML templates and assets. The `dist/` folder is what gets served/deployed.
