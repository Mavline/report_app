# Excel Report Manager Prototype

## What this project demonstrates

A React-based prototype for importing an Excel workbook, selecting worksheet data, grouping rows, ordering fields, and preparing a structured report output. The project demonstrates spreadsheet workflow software for internal reporting and operational review.

## Use case

This type of tool is useful when a team receives spreadsheet-based operational data and needs a repeatable interface for reviewing, restructuring, and exporting report-ready results without rebuilding the workflow manually each time.

## Features

- Excel file upload and parsing.
- Worksheet selection.
- Field selection and ordering.
- Grouping logic for report preparation.
- Preview and export-oriented workflow.
- Reset flow for repeated processing.

## Technical stack

- Frontend: React 18, TypeScript, Create React App / react-app-rewired.
- UI: Radix UI primitives, Lucide icons, custom CSS.
- Spreadsheet/data: xlsx, ExcelJS, JSZip, XML parsing utilities.
- Export: file-saver.

## Architecture

The browser UI accepts an Excel workbook, parses workbook and worksheet data on the client side, lets the user select fields and grouping settings, then prepares structured output for reporting or downstream review.

## Screenshots

A deployed Replit screenshot is stored in the acty.dev proof asset inventory as `alereport-manager-excel-report.png`.

## How to run locally

```bash
npm install
npm start
```

Build:

```bash
npm run build
```

## Notes

This repository is a sanitized proof-of-work edition. It does not include private client data, production credentials, internal datasets, or confidential business logic.
