# Manager Excel Report

## Overview
A web-based tool for merging Excel report data from multiple sheets. Used for supply chain/logistics management (A.L. Electronics / Novocure). Users upload an Excel file, select two sheets, map fields via drag-and-drop, then merge and download a standardized report.

## Tech Stack
- React 18 + TypeScript (CRA with react-app-rewired)
- SheetJS (xlsx 0.18.5) + ExcelJS for Excel file processing
- Tailwind CSS + Radix UI for styling
- Static deployment (build folder served)

## Architecture
- Single Page Application, main logic in `src/App.tsx`
- Client-side Excel parsing and merging
- No backend — all processing happens in the browser

## Key Files
- `src/App.tsx` — Main application component (all business logic)
- `src/components/ui/` — Reusable UI components (Input, Button, etc.)
- `package.json` — Dependencies and build scripts

## Important Notes
- Date formatting uses a deterministic function (`formatDateDeterministic`) instead of `toLocaleDateString` to avoid browser-dependent ICU/locale changes breaking date key matching
- Excel serial dates are converted via `excelSerialToDate` using UTC to avoid timezone drift
- The merge button shows an alert when no matching data is found

## Deployment
- Static deployment: `build/` directory
- Dev server: `npx serve -s build -l 5000`
