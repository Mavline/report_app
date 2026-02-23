# Project Brief

## Project
`report_app` is a browser-based Excel report merger for Novocure / A.L. Electronics workflows.

## Primary Goal
Allow a user to upload an Excel workbook, map columns from two sheets, merge rows by part number, and export a standardized merged report.

## Core Value
Reduce manual Excel reconciliation work while preserving date-based quantity columns and delivery/balance calculations.

## Current Risk Area
Date column labels and date formatting across Excel files, SheetJS parsing, UI drag-and-drop labels, and merge key lookup.

## Scope Notes
- Frontend-only app (no backend)
- Main business logic is concentrated in `src/App.tsx`
- Static deployment (Replit/deployed build)

