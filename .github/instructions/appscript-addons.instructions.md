---
applyTo: "Gsheets addons/**,Gslides addons/**"
description: "Use when implementing or fixing Google Apps Script addon features in Gsheets addons or Gslides addons, especially config-sheet contracts and google.script.run client-server flows."
---

# Apps Script Addons Instructions

## Mission

Implement and fix features in the Sheets/Slides addon pair without breaking the shared config contract.

## Required Checks Before Editing

- Identify whether the change touches only one addon or both.
- If data exchanged via config rows or settings JSON changes, update both producers and consumers.
- Keep response object fields stable unless explicitly migrating all callers.

## Implementation Rules

- Trace UI event handler in `*.html` to corresponding server function in `*.gs` before editing.
- Preserve existing `configSheetId` and `Configs` tab usage.
- Keep edits small and local; avoid broad refactors unless requested.
- Retain manual-debug visibility with `Logger.log` and `console.log` at key flow boundaries.

## Contract-Safe Change Pattern

1. Add new fields as optional first.
2. Keep old fields readable during transition.
3. Update UI rendering and server parsing together.
4. Guard missing values with defaults.

## Done Criteria

- No broken `google.script.run` calls.
- Config setup still links and validates sheet access.
- Config loading and sync/export actions still return actionable success/error messages.
