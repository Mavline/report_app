# Active Context

## Date
2026-03-12

## Current Session Goals
- Review recurrence of date-related merge mismatch
- Improve universality of date normalization for `Qty-by-date`
- Reduce reliance on exact string shape of date labels during matching and sorting

## What Was Done In This Session
- Reviewed the existing `Qty-by-date` normalization path in `src/App.tsx`
- Identified that previous matching was still based on handcrafted string variants and exact display-label shape
- Replaced narrow date normalization with canonical date fingerprints for both dragged labels and actual row keys
- Expanded matching to tolerate punctuation, multiple separators, Unicode spaces/NBSP, month token variants, 2-digit vs 4-digit years, and numeric Excel-serial-like date strings
- Replaced ambiguous `new Date(string)` sorting for mapped/merged dates with canonical date sort values
- Verified by local production build on 2026-03-12: build succeeds
- Observed a follow-up issue after the first hardening pass: March 26 reappeared, but April dates could still be mislabeled/missed
- Identified the deeper cause: `Qty-by-date` mapping still derived its canonical display date from the dragged header string, which can be ambiguous if Excel switches default header rendering (for example numeric month/day display)
- Added sheet header metadata so canonical mapped dates now come from the actual Excel header cell, using `cell.v`/serial date when available instead of guessing only from `cell.w`
- Updated mapped source-field resolution to prefer exact sheet metadata for dragged fields during merge
- Verified the current fix on the correct production-relevant workbook after a temporary detour through an outdated test file

## Key Findings (Date Issue)
- Previous fixes solved exact known variants but still depended on a finite list of string rewrites.
- A single newly added date can still fail when its header formatting differs in punctuation, spacing, separator style, or year representation from previous dates.
- The robust approach is to canonicalize both sides of the comparison and match by canonical date identity, not by precomputed textual variants alone.
- Even canonical string matching is not enough for ambiguous numeric date headers like `4/2/2026`; the mapped date label must be derived from the underlying Excel cell value, not only from the displayed string.
- The production issue itself was real; the mistaken part was only a later verification pass that used an outdated workbook and briefly distorted the interpretation of side effects.

## Current Code Anchors
- `src/App.tsx` `getValueByDateKey` (robust date-key matching)
- `src/App.tsx` `getCellDisplayValue` (prefers `cell.w`)
- `src/App.tsx` `getSheetHeaders` / `filterAndFormatHeaders` (header display generation)
- `src/App.tsx` `mergeTables` no-match alert path
- `src/App.tsx` `buildDateFingerprints` (canonical date fingerprints for label/key matching)
- `src/App.tsx` `getDateSortValue` (stable date ordering without locale-dependent string parsing)
- `src/App.tsx` `getNormalizedDateFromHeaderCell` / `getSheetHeaderDetails` (canonical header-date metadata from actual Excel cells)

## Open Questions
- Does the problematic workbook use one of the newly covered variants (comma, slash, dot, NBSP, 4-digit year, numeric-like date label)?
- Should old debug logs remain in deployed code, or be reduced after confirming the fix on Replit?
- After switching mapped date normalization to header-cell metadata, can debug logs now be reduced or removed?

## Next Recommended Actions
1. Commit and push the validated fix
2. Redeploy the fresh production build to Replit/site
3. Optionally reduce debug logging after deployment confirmation
