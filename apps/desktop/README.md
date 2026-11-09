# BetterOffice desktop editors

The Docs, Sheets, and Slides editor UI: the same React editors shipped by
`apps/demo`, including their toolbars, menus, dialogs, status bars, sheet tabs,
slide thumbnails, editing commands, and collaboration presence, behind a start
screen that opens without a sample document. Capy keeps it as the bundle the
browser end-to-end tests serve (`apps/desktop/dist`); its macOS shell was
removed with the native viewer.

Run the UI in a browser for development:

```sh
bun run --filter @betteroffice/desktop dev
```

Use `?format=xlsx` or `?format=pptx` to preview the other apps. Browser previews
use browser file selection and downloads.

Validation:

```sh
bun run --filter @betteroffice/desktop typecheck
bun run --filter @betteroffice/desktop test
bun apps/desktop/tests/smoke.ts
```
