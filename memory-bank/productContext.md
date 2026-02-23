# Product Context

## Users
- Internal/corporate users (supply chain / logistics workflow)
- Non-developer operators using uploaded Excel files from external/internal sources

## Typical Workflow
1. Upload an Excel workbook
2. Select left/right sheets
3. Choose header rows / fields
4. Drag source fields to template fields (including multiple date columns)
5. Run merge
6. Preview results
7. Download merged Excel

## User Pain Points Observed
- Merge appears to do nothing when mappings are filled
- "Not matching" style behavior when date columns visually look correct but keys do not match internally
- Date formats vary by file / Excel formatting (`Sep`, `Sept`, leading zeros, separators)

## Why Date Issues Hurt
The app relies on date-labeled quantity columns (`Qty-by-date`). If date labels do not match parsed row keys exactly, quantities are not found, and merge output can be empty.

## Current User-Reported Context (2026-02-23)
- Deployed Replit app reportedly stopped responding on Merge for a corporate user
- User recalled previous similar date-format issue from months earlier
- Remote agent attempted fixes; concern was that behavior worsened with "not matching" warnings before final fixes

