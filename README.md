# Client Comparison Tool — PSU + Local Code Matching + Force

## Workbook output order
All valid client sheets are kept first in their existing order. Three utility sheets are then added in this exact order:
1. `PSU` — enter a client worksheet name in B2; returns `Matched_CL_ID`, `PTY_LOCAL_CODE`, and `Source_Client_Sheet` from rows across all client sheets where `Matched_Sheet` equals that client name. Rows are included even when `PTY_LOCAL_CODE` is blank.
2. `Local Code Matching` — enter a client worksheet name in B2; returns the source client sheet name, the actual source-row `CL_ID`, and its `PTY_LOCAL_CODE`.
3. `Force` — permanent three-column mapping register: `Client_Unique_ID`, `PTY_LOCAL_CODE`, and `Source_Client_Sheet`.

## Force mapping register
- `Client_Unique_ID`: the unique ID to look up against `Matched_CL_ID` in all client comparison sheets.
- `PTY_LOCAL_CODE`: the local/TD code to populate beside the matching client row.
- `Source_Client_Sheet`: the client worksheet from which the mapping was obtained, for traceability.
- Keep existing entries and add new mappings on new rows. The tool carries all existing rows with a Client_Unique_ID forward, including rows where PTY_LOCAL_CODE is blank.
- Re-running comparison does not clear the Force register. Existing client-sheet PTY_LOCAL_CODE values are retained when there is no nonblank mapping available.
- The mapping is applied globally across all client sheets. The source sheet is tracking information and does not restrict where the ID is matched.

## Comparison behavior
- Compares valid client worksheets in reverse order: the last client sheet is the reference; earlier sheets compare against sheets to their right.
- Excludes `PSU`, `Local Code Matching`, `Force`, and legacy utility sheet names from business comparison.
- Adds/retains `Matched_Sheet`, `Match_Type`, `Matched_CL_ID`, and `PTY_LOCAL_CODE` on client sheets.
- Uses the global mapping register in `Force` for every client worksheet. It looks for a nonblank code for the matching ID.

## Extraction details
- `PSU`: filters all client sheets by `Matched_Sheet` equal to the entered client name and returns `Matched_CL_ID`, `PTY_LOCAL_CODE`, and `Source_Client_Sheet`, including blank-code rows.
- `Local Code Matching`: filters all client sheets by `Matched_Sheet` equal to the entered client name and returns the source sheet name, actual source-row `CL_ID`, and `PTY_LOCAL_CODE`.
- Enter the exact worksheet name in cell B2. Dynamic extraction requires Microsoft Excel 365 or Excel for the web (`LET`, `VSTACK`, `FILTER`, and `CHOOSECOLS`).

## Run locally / GitHub Pages
Open `index.html` in a browser or publish the files in the repository root using GitHub Pages. SheetJS is loaded from jsDelivr. Workbook processing happens locally in the browser; client data is not uploaded by this app.

## Files
- `index.html` — application interface
- `app.js` — comparison, mapping, and extraction logic
- `style.css` — styling
