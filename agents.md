# sheets2slides Agent Guide (Apps Script Focus)

This workspace contains multiple stacks, but active development focus is the Google Apps Script addon code:
- `Gsheets addons/` (config export addon)
- `Gslides addons/` (image sync addon)

Use this guide for implementing features and fixing bugs in those two directories.

## Scope For Agents

- Prioritize changes in `Gsheets addons/` and `Gslides addons/`.
- Treat `terraform/` and `pub-sub/` as out of scope unless explicitly requested.
- Keep changes minimal and behavior-preserving unless the request asks for a workflow change.

## Architecture Snapshot

- Sheets addon backend script: `Gsheets addons/exportconfig.gs`
- Sheets addon sidebar UI: `Gsheets addons/exportconfigfront.html`
- Slides addon backend script: `Gslides addons/imagesync.gs`
- Slides addon sidebar UI: `Gslides addons/imagesyncsidebar.html`
- Legacy high-level context: `agentsold.md`

## How The Two Addons Communicate

The shared contract is a centralized config spreadsheet plus user property state.

- Shared user property key: `configSheetId`
- Required tab in config spreadsheet: `Configs`
- Key server functions on both sides:
  - `getConfigSheetId()`
  - `setConfigSheetId(sheetId)`
  - `checkConfigSheet()`
- Sheets addon writes/updates config rows.
- Slides addon reads config rows and uses them to sync image replacements.

Current `Configs` header schema (from Sheets addon):
1. `Config ID`
2. `Name`
3. `Spreadsheet ID`
4. `Spreadsheet Name`
5. `Sheet Tab`
6. `Cell Range`
7. `Created`
8. `Updated`
9. `Settings JSON`

## Editing Conventions (Critical)

- Keep `google.script.run` client calls aligned with server function names and return shapes.
- If a server response object changes, update all sidebar handlers that read its fields.
- Preserve existing config shape fields used in both addons (for example settings/filter/png fields).
- Prefer additive migration-safe changes over breaking schema rewrites.
- Maintain existing logging style (`Logger.log` in `.gs`, `console.log` in `.html`) when troubleshooting.

## Feature Work Workflow

1. Locate entry point from UI handler in the relevant `*.html` file.
2. Trace matching server function in the relevant `*.gs` file.
3. Identify any cross-addon contract impact (`Configs` row schema, `configSheetId`, settings JSON).
4. Implement minimal code changes in both client/server sides when needed.
5. Validate by checking:
   - Config setup view can link/check sheet.
   - Config list loads correctly.
   - Action flow returns expected success/error payloads.

## Common Pitfalls

- Function name mismatches between sidebar JS and Apps Script server code.
- Changing settings JSON shape without backward compatibility.
- Assuming file names equal Apps Script project HTML names; verify names referenced by `HtmlService.createHtmlOutputFromFile(...)`.
- Silent regressions when changing `Configs` columns or order.

## Source-Of-Truth Docs

- Existing project context: `agentsold.md`
- Planning docs: `IMPLEMENTATION_PLAN.md`, `IMPLEMENTATION_PLAN_BIGSCALE.md`

Link to these docs instead of duplicating long explanations in future instruction updates.
