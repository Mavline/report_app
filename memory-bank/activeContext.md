# Active Context

## Date
2026-02-23

## Current Session Goals
- Sync local repo with remote changes from Replit/deployed fixes
- Inspect historical date-format commits to understand prior fixes
- Create project `AGENTS.md` and initialize Cline-style `memory-bank`

## What Was Done In This Session
- Confirmed local working tree was clean before sync
- Pulled remote updates with fast-forward: `df777e3 -> f75ae12`
- Reviewed recent commits and date-related history
- Inspected current `src/App.tsx` date logic (`getValueByDateKey`, `getCellDisplayValue`, header formatting functions, merge no-match alert)
- Added frontend-side debug logs for:
  - sheet parsing (`[sheetSelection]`)
  - drag-and-drop mapping (`[mapping]`)
  - merge pipeline (`[merge] start`, `[merge] diagnostics`)
- Reproduced runtime symptom path: `Merge` click works, but merge ends with `No matching data found`
- Identified concrete blocker from logs: SO balance field reads as `undefined` because SheetJS key includes spaces, while merge code used trimmed hardcoded key
- Applied minimal fix preserving hardcoded semantics by resolving actual row keys via whitespace-normalized header matching for fixed keys (`ALE PN`, `מקט`, `יתרה לאספקה`, `תאריך מובטח`)

## Key Findings (Date Issue)
- The Aug 19, 2025 fix (`df777e3`) improved date label matching (`Sep/Sept`, leading zero) but did not fully align UI header labels with SheetJS row keys.
- The 2026-02-23 fixes addressed the deeper mismatch by using `cell.w` for header display values (same source used by SheetJS formatting) with deterministic fallback.
- Current code also broadens date-key lookup variants and shows a user alert when merge yields no matching data.
- Current runtime failure (during this session) was not the button itself: merge executed and failed because balance lookup used a mismatched hardcoded SO key (`'יתרה לאספקה'` vs actual key with spaces).

## Current Code Anchors
- `src/App.tsx` `getValueByDateKey` (robust date-key matching)
- `src/App.tsx` `getCellDisplayValue` (prefers `cell.w`)
- `src/App.tsx` `getSheetHeaders` / `filterAndFormatHeaders` (header display generation)
- `src/App.tsx` `mergeTables` no-match alert path

## Open Questions
- Is the currently deployed Replit build already using commit `f75ae12` artifacts, or is deployment lagging behind repo state?
- Do user files contain additional header variants not yet covered (dots, localized month names, hidden spaces)?
- Is there any UI condition that still prevents Merge click handler execution (separate from no-match logic)?
- After whitespace-tolerant hardcoded-key resolution fix, does merge now produce rows for the problematic workbook?

## Next Recommended Actions
1. Reproduce merge locally with a real problematic workbook (if available)
2. Add lightweight debug logging/toggle around parsed headers vs row keys (only if issue persists)
3. Consider extracting date/header matching helpers out of `src/App.tsx` for targeted tests
