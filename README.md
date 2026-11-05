# Client Comparison Tool

Browser-based Excel comparison tool. Workbook processing runs locally in the browser.

## Output workbook

- Original client worksheets remain in their original order.
- `PSU`: enter a client worksheet name in **B2**. Returns `Matched_CL_ID`, `PTY_LOCAL_CODE`, and `Source_Client_Sheet`. Rows are retained even when `PTY_LOCAL_CODE` is blank.
- `Force`: enter a client worksheet name in **B2**. Returns `Source_Client_Sheet`, the source row's actual `CL_ID`, and `PTY_LOCAL_CODE`.
- `Local Code Logic`: persistent three-column mapping register: `Client_Unique_ID`, `PTY_LOCAL_CODE`, `Source_Client_Sheet`.
- `_ExtractData`: hidden helper worksheet used by the two extraction sheets.

## Comparison

- Preserves the existing reverse-order comparison: earlier client sheets compare against sheets to their right; the last client sheet is the reference.
- Preserves the existing banner/address normalization and matching logic.
- Excludes `PSU`, `Force`, `Local Code Logic`, and recognized legacy utility sheets from client comparisons.
- Mapping IDs are matched globally against `Matched_CL_ID`. `Source_Client_Sheet` is retained as traceability information.
- Existing local codes are retained when no nonblank mapped code is found.
- Existing mapping rows are carried forward when the previously generated workbook is used as the next input. Add new mappings below existing entries; existing entries are not overwritten.

## Files

- `index.html` — application page
- `app.js` — workbook parsing, comparison, extraction, and download
- `style.css` — interface styling

Publish these files in the GitHub Pages repository root. SheetJS is loaded from jsDelivr. Workbook contents are processed locally in the browser.
