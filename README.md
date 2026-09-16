# Client Comparison Tool

Browser-based Excel comparison tool.

Logic:
- Sheet 1 is oldest and has no Matched_Sheet column.
- Sheet 2 compares with Sheet 1.
- Sheet 3 compares with Sheet 1 first, then Sheet 2.
- Each later sheet compares with all older sheets in workbook order.
- First matching older sheet wins.
- Matching uses CL_STORE_NAME + CL_PA_LINE1.
- No-match rows stay blank.
- All sheets can be uploaded at once, or new sheets can be added to the workbook later.

GitHub Pages can host this static site. Do NOT commit client Excel files to the repository.

The app uses SheetJS from jsDelivr in the browser.
