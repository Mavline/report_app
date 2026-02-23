# Change Log (Memory Bank)

## 2026-02-23 (local session)
- Ran `git pull --ff-only` and synced local `main` from `df777e3` to `f75ae12`.
- Pulled commits included:
  - `d06f551` Improve date handling and add merge feedback to the reporting tool
  - `b327e33` Improve date matching and handling for uploaded files
  - `2b29402` Published your App
  - `a825460` Published your App
  - `f75ae12` Published your App
- Verified current code in `src/App.tsx` now:
  - uses deterministic date formatting (`formatDateDeterministic`, `excelSerialToDate`)
  - prefers `cell.w` via `getCellDisplayValue` for header labels
  - expands `getValueByDateKey` matching variants
  - alerts when merge produces zero matching rows
- Added runtime diagnostics in `src/App.tsx` for:
  - `[sheetSelection]` (header row index, parsed rows, first row keys)
  - `[mapping]` (field mapping assignments and state snapshots)
  - `[merge] start` and `[merge] diagnostics` (counters and sample failures)
- User-provided console logs showed:
  - `Merge` handler executes successfully (button click issue was symptom masking empty result)
  - SO `balance` samples are `undefined` in diagnostics
  - right sheet keys include spaced variant of balance header (`" יתרה לאספקה           "`)
- Applied minimal fix in `src/App.tsx`:
  - keep hardcoded logical field names
  - resolve actual SheetJS row keys by whitespace-normalized header matching (`resolveHeaderKey`)
  - use resolved keys for `ALE PN`, `מקט`, `יתרה לאספקה`, `תאריך מובטח` during merge

## Historical Date-Related Commit Notes

### 2025-08-19 `df777e3` Adjusted error of reading date-month from cells
- Replaced direct normalized key lookup with `getValueByDateKey(row, qtyField)`.
- Expanded `normalizeDate` and added alternative matching for `Sep/Sept` and day leading zero.
- Limitation: still a partial fix because UI header labels and SheetJS row keys could be generated differently.

### 2024-12-03 `f4741e8` Improved dates visible and format
- Focused on exported Excel formatting for `Delivery-Requested` / `Delivery-Expected`.
- Converted string fields to `Date` objects and applied ExcelJS `numFmt`.
- Also adjusted date mapping normalization in drop handling.

### 2024-12-03 `0269481` Improved dates visible and format
- Largely rolled back/simplified parts of `f4741e8` export-date conversion.
- Reverted explicit ExcelJS date column formatting and `new Date(...)` conversion for exported rows.
- Kept improvements around mapped date normalization in field drop logic.

### 2024-11-17 `792634f` Dates formatted
- Introduced `toLocaleDateString('en-GB', ...)` formatting for `Delivery-Requested` converted from Excel serial.
- This improved display but introduced dependency on browser locale/ICU behavior for generated date strings.

### 2024-11-17 `346c079` Working version with not date format in 1 column
- Temporarily removed `formatDate` usage for `Delivery-Requested` in merged rows (raw value passed through).
- Indicates active instability around date formatting during early merge implementation.

### 2024-11-17 `d909512` Splitted dates, without balance
- Major merge logic restructuring introducing per-date row splitting behavior.
- Early stage of date-column merge behavior before later formatting/matching stabilizations.

## 2026-02-23 (memory infrastructure)
- Added project-level `AGENTS.md` with rules for maintaining memory bank.
- Initialized `memory-bank/` with 7 markdown files (Cline-style + change log).
