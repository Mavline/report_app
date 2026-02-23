# Tech Context

## Stack
- React 18 + TypeScript
- Create React App (via `react-app-rewired`)
- SheetJS (`xlsx`) for reading workbook/sheets
- ExcelJS for export generation
- Radix UI + Tailwind utility usage

## Key Files
- `src/App.tsx` (main business logic, date parsing/matching, merge, export)
- `src/components/ui/*` (UI primitives)
- `package.json` (scripts/dependencies)
- `.replit` / `replit.md` (Replit deployment/runtime metadata)

## Commands
- `npm start`
- `npm run build`
- `npm test`

## Repository / Branch
- Remote: `origin https://github.com/Mavline/report_app.git`
- Main branch: `main`

## Current Local State (after sync on 2026-02-23)
- Pulled `origin/main` fast-forward from `df777e3` to `f75ae12`
- Recent commits include date fixes and Replit publish commits

## Important Libraries / Behaviors
- SheetJS `sheet_to_json` header behavior can depend on cell formatted text (`cell.w`)
- Excel serial date conversion must avoid local timezone side effects
- ExcelJS export date formatting can be brittle if string dates are converted to `Date` inconsistently

