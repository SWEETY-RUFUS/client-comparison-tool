# Client Comparison Tool

Browser-based Excel comparison tool.

## Matching logic

- Sheet order is preserved.
- Sheet 2 checks Sheet 1 first.
- Sheet 3 checks Sheet 1, then Sheet 2.
- Each later sheet checks all older valid client sheets in order.
- The first/oldest match wins.
- Primary fields:
  - `CL_STORE_NAME`
  - `CL_PA_LINE1`
- Exact Banner + Address is an exact match.
- Normalized address variations such as STREET/ST, AVENUE/AVE, SUITE/STE and common punctuation/spacing are handled.
- Banner variations can match when the address is the same/strongly equivalent.
- Address variations can match when the banner is the same/strongly equivalent.
- `Matched_Sheet` is always added to valid comparison sheets.
- `Match_Type` records `Exact`, `Address + Banner variation`, or `Banner + Address variation`.
- Sheets missing either required column are left unchanged and shown as warnings.
- A newly uploaded workbook resets the previous workbook completely.

## GitHub Pages

Upload `index.html`, `style.css`, and `app.js` to a GitHub repository and publish the repository with GitHub Pages.

The Excel file is processed in the browser and is not uploaded to your GitHub repository by this app. Do not commit client Excel files to GitHub.


## Performance update

Version 3 uses lookup indexes instead of scanning every older row for every current row.
This prevents the browser from becoming unresponsive on larger workbooks.

The comparison still follows the original sheet order and first/oldest-match-wins rule.
