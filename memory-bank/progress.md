# Progress

## Completed
- 2026-03-12: Added a UI-only `Pro` header filter so only current-year date columns are shown to the operator
- 2026-02-23: Local repo synced to latest remote (`f75ae12`) via fast-forward pull
- 2026-02-23: Verified current code includes:
  - deterministic date formatting helpers
  - `cell.w`-based header display fallback logic
  - expanded `getValueByDateKey` matching
  - "No matching data found" merge feedback alert
- 2026-02-23: Reviewed historical date-format commit chain (2024-11 through 2026-02)
- 2026-02-23: Initialized memory bank structure and AGENTS instructions
- 2026-02-23: Added runtime diagnostics for mapping and merge pipeline in `src/App.tsx`
- 2026-02-23: Identified SO balance key mismatch due to SheetJS-preserved header spaces (`יתרה לאספקה`) causing all rows to be filtered out
- 2026-02-23: Implemented whitespace-tolerant resolution for fixed hardcoded merge keys (without changing hardcoded workflow)
- 2026-03-12: Reworked `Qty-by-date` matching from ad hoc string variants to canonical date fingerprints
- 2026-03-12: Replaced `new Date(string)` sort usage for mapped/merged date labels with canonical sort values
- 2026-03-12: Verified local `npm run build` succeeds after the date-normalization changes
- 2026-03-12: Added sheet header metadata so mapped date labels come from actual Excel header cells, not only displayed header strings
- 2026-03-12: Confirmed that the production issue was real and the fix was validated against the correct workbook; only an intermediate side-effect check used an outdated test file

## In Progress
- Prepare the validated fix for commit/push and deployment

## Pending
- Confirm that the `Pro` field panel now hides pre-2026 date columns in the deployed UI
- Confirm the deployed Replit/site is running the fresh build
- Decide whether to add regression tests for date header matching variants
- Potentially refactor `src/App.tsx` helper functions into testable utilities

## Risks
- Recurring date regressions due to mixed formatting systems (Excel formatting vs JS formatting)
- Single-file architecture (`src/App.tsx`) increases accidental coupling during fixes
- Replit publish commits may obscure which code change actually fixed behavior

## Notes
- Local `npm run build` was run successfully in this session
- `node_modules` exists locally, so local build/run should be possible if needed
