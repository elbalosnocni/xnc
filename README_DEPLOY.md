# Foreign Employee Management V2.1

## V2.1 highlights
- Master Excel preview before import.
- Row-level validation for EmployeeMaster and MovementPlan.
- Import is blocked when preview contains validation errors.
- One workbook supports EmployeeMaster + MovementPlan.
- Admin batch import for Entry/Exit and Domestic Travel.
- Employee timeline with filters: All / Entry-Exit / Travel.
- Entry/Exit stores PlannedDateTime, ActualDateTime and TripID.
- Google Drive document preview remains supported for Passport, TRC, Work Permit and Contract.
- VI / EN / ZH translation remains compatible with existing UI.

## Deploy
1. Open the Apps Script project bound to the target Google Sheet.
2. Replace `Code.gs` with the V2.1 `Code.gs`.
3. Replace the web UI file with `index_fixed.html` and rename it to `index` if your Apps Script HTML file is named `index`.
4. Run `setupSystem()` once from Apps Script.
5. Authorize Spreadsheet / Drive permissions.
6. Deploy a new Web App version.
7. Confirm the frontend API URL/config matches the deployed Web App.

## Master Excel
Use `Download Template` from Foreign Employees. The workbook contains:
- `EmployeeMaster`: employee + Passport + TRC + Work Permit + Contract + Appendix.
- `MovementPlan`: Entry/Exit + Domestic Travel.

Workflow:
1. Admin uploads workbook.
2. System parses both sheets.
3. Preview shows total / valid / error / duplicate counts.
4. Filter errors and fix the source Excel if needed.
5. Only when there are no validation errors can Admin click Confirm & Import.
6. Employee data is upserted first, then MovementPlan is imported.

## Important
- Do not remove existing Google Sheets data before upgrading.
- `EntryExit` now uses 13 columns: LogID, EmployeeID, Type, EventDate, PortName, FlightNo, Destination, Purpose, CreatedAt, CreatedBy, PlannedDateTime, ActualDateTime, TripID.
- If an older EntryExit sheet has fewer columns, run `setupSystem()` once; it adds the missing columns without deleting rows.
