# Ethan Excel v19 — Functional Upgrade Notes

Date: 2026-11-08
Production app: https://office.ethandigitalacademy.org
Repository: abayomipeteroyenuga-commits/Ethanoffice
Branch: main

## Reference research
The feature baseline was reviewed against Microsoft Excel for the web documentation and the public feature descriptions for Google Sheets and other browser spreadsheet products. The comparison focused on interactive cell selection, formula entry, calculation feedback, data handling and common spreadsheet workflows.

Reference:
- Microsoft Excel for the web feature availability: https://learn.microsoft.com/en-us/office365/servicedescriptions/office-online-service-description/excel-online
- Microsoft Excel for the web service description: https://learn.microsoft.com/office365/servicedescriptions/office-online-service-description/office-online-service-description
- Microsoft Excel browser behaviour: https://support.microsoft.com/en-us/excel/differences-between-using-a-workbook-in-the-browser-and-in-excel

## v19 changes
- Replaced the fragile cell activation/render path with a consistent active-cell indicator and address display.
- Clicking a cell activates it and loads its raw value/formula into the formula bar.
- Added keyboard movement with Enter, Tab, Shift+Tab, Arrow Up and Arrow Down.
- Added Shift-click rectangular range selection.
- Status bar aggregates the selected cell/range: Sum, Average and Count.
- Formula engine supports SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, ROUND and arithmetic with cell references.
- Formula references and ranges work across columns A:T.
- Formula recalculation refreshes dependent visible cells after an edit.
- AutoSum inserts a SUM formula for the contiguous numeric/data block above the active cell.
- Formula library provides examples and a one-click insertion action.
- Improved CSV parsing for quoted fields and commas/newlines inside quoted values.
- CSV export quotes values containing commas, quotes or line breaks.
- Preserved existing worksheet formatting, validation metadata, freeze-row state, autosave and protection behaviour.
- Added clearer visual feedback for active and range-selected cells.

## Existing capabilities retained
- 60 rows × 20 columns (A:T).
- Ribbon tabs for Home, Insert, Formulas, Data, Page Layout, Review, View and File.
- Cell formatting, currency/percent styles, fill down, clear, row/column insertion, sorting, filtering, find/replace, duplicate removal, dropdown validation, charts, freeze row, sheet protection, CSV import/export and print/PDF workflow.

## Known next-phase work
- Full multi-sheet workbook tabs and sheet rename/duplicate/reorder.
- Formula autocomplete and a much broader function library.
- Robust range copy/paste, fill handle and drag-to-select.
- Real filter dropdowns per column and sortable table headers.
- Conditional formatting rules with saved rule definitions (not one-cell visual toggles).
- Native XLSX import/export that preserves formulas, styles and charts.
- Charts embedded and editable inside the worksheet.
- Formula dependency graph, circular-reference handling UX and formula error explanations.
- Multi-user cloud collaboration and workbook version history.

## Verification status
- GitHub main branch updated.
- Vercel production deployment reports READY for the v19 commit.
- The custom domain remains attached and verified in the Vercel project configuration.
- A full browser click-through was not available in this execution, so real user interaction should be smoke-tested after refresh before treating every control as fully verified.
