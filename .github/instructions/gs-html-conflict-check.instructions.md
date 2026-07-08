---
name: "GS HTML Conflict Check"
applyTo: "Gsheets addons/**/*.gs,Gslides addons/**/*.gs,Gsheets addons/**/*.html,Gslides addons/**/*.html"
description: "Use when editing Apps Script .gs or addon sidebar .html files in Gsheets addons or Gslides addons. Ensure server-client compatibility and no google.script.run handler or response-shape conflicts are introduced."
---

# GS to HTML Compatibility Check

When changing any `.gs` or `.html` file under `Gsheets addons/` or `Gslides addons/`, verify the related server/client code still matches.

## Required Checks

- Find corresponding `google.script.run` calls in the related `.html` file and confirm server function names still match.
- If a server response object changes, update all `.html` consumers reading those fields.
- If function parameters or required data change in `.gs`, update `.html` call sites and validation paths.
- Keep shared config contract fields backward compatible unless a coordinated migration is explicitly requested.

## Done Criteria

- No `.gs`/`.html` function-name mismatches.
- No response-shape mismatches between server and sidebar code.
- The related sidebar flow still returns actionable success/error messaging.
