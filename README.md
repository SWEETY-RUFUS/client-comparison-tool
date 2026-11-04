# Client Comparison Tool — Reverse Order + CL_ID + Global Force Mapping + Client Extract

## What it does
- Compares valid client worksheets in reverse order: the last client sheet is the reference; earlier sheets compare against sheets to their right.
- Excludes `Force` and `Matched_ID_Extract` from business comparison.
- Adds `Matched_Sheet`, `Match_Type`, `Matched_CL_ID`, and `PTY_LOCAL_CODE` to client worksheets.
- Uses one global `Force` mapping table for every client worksheet.
- Preserves existing Force mapping rows when the workbook is uploaded again. Add new mappings below old rows; do not clear the previous mappings.
- Adds one reusable `Matched_ID_Extract` sheet. Enter any exact client worksheet name in cell `B2` to list its `Matched_CL_ID` and `PTY_LOCAL_CODE` values. Blank TD codes are shown as blank.

## Force mapping
The `Force` sheet has only two columns: `Client_Unique_ID` and `PTY_LOCAL_CODE`. The `Matched_CL_ID` in every client sheet is looked up against this shared list. The first mapping for a given ID is used. Keep all earlier mapping rows and append new IDs below them so previous allocations remain available.

## Client-wise matched ID extract
In `Matched_ID_Extract`, type the exact worksheet name in `B2` (for example `SGWS`). The dynamic result displays the corresponding `Matched_CL_ID` and `PTY_LOCAL_CODE` for every row across the client sheets where `Matched_Sheet` equals the entered name. It includes rows whose PTY code is blank. This automatic extraction requires Microsoft Excel 365 or Excel for the web (dynamic array functions such as `LET`, `VSTACK`, `FILTER`, and `CHOOSECOLS`).

## Run locally / GitHub Pages
Open `index.html` in a browser or publish the files in the repository root using GitHub Pages. SheetJS is loaded from jsDelivr. Workbook processing happens locally in the browser; client data is not uploaded by this app.

## Files
- `index.html` — application interface
- `app.js` — comparison, Force mapping, and extraction logic
- `style.css` — styling
