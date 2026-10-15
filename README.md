# HRMS Timesheet (GitHub Pages)

A lightweight personal timesheet app. No backend/server required.

## Features
- Optional timesheet fields; Date is required
- Automatic month, serial number, and duration calculation
- Manage Code, Product, Business Name, and Status dropdown values
- Search and filter records
- Export filtered, selected-day, or selected-month records to `.xlsx`
- Import CSV/Excel with success/failure counts and row-level reasons
- JSON backup/restore
- Confirmation before deleting entries or clearing all data

## Run locally
Open `index.html` in a modern browser. Excel import/export uses SheetJS from jsDelivr, so internet is needed for those features.

## Publish on GitHub Pages
1. Create a GitHub repository (e.g. `hrms-timesheet`).
2. Upload `index.html` and this `README.md` to the repository root.
3. Open **Settings → Pages**.
4. Under Build and deployment, choose **Deploy from a branch**.
5. Select `main` and `/(root)`, then Save.
6. Open the URL GitHub provides.

## Important data storage note
GitHub Pages hosts static files only. This version saves entries in the current browser's `localStorage`; it does not sync between devices/browsers and clearing browser data may remove entries. Use **Download full backup (JSON)** regularly. For multi-device sync or secure centralized storage, connect a backend/database or Google Sheets API later.

## Import file format
The first worksheet / CSV should have headers such as:
`Date, Code, Source, Product, Transaction Start Time, Transaction End time, CONTACT NO, Business Name, Issue/Request, Ticket No, Status, Remarks`

Only Date is required. Missing optional fields are allowed. Import appends records and automatically adds imported Code/Product/Business/Status values to dropdown master data.
