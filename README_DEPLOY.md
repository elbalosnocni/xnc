# Foreign Employee Management V2 - Unified Master Excel

## Files
- `index_fixed.html`: frontend. Keep as `index.html` in Google Apps Script.
- `Code.gs`: backend. Replace the old `Code.gs` with this file.

## Main changes
1. Unified Master Excel workbook:
   - `EmployeeMaster`: employee + Passport + TRC + Work Permit + Contract + Appendix.
   - `MovementPlan`: Entry/Exit + Domestic Travel in one sheet.
2. Backend actions:
   - `BATCH_ADD_ENTRY_EXIT`
   - `BATCH_ADD_DOMESTIC_TRAVEL`
   - `BATCH_IMPORT_MOVEMENT`
3. EntryExit schema is backward compatible in the first 9 columns and adds:
   - `CreatedBy`
   - `PlannedDateTime`
   - `ActualDateTime`
   - `TripID`
4. `setupSystem()` updates the EntryExit header without deleting existing rows.
5. Frontend supports key-based i18n (`data-i18n`) while retaining legacy translation for existing UI.
6. Google Drive preview URLs are corrected.
7. Admin/HR Admin can batch import movement; only ADMIN can import the full Master Excel.

## Deploy
1. Open Apps Script project.
2. Replace `Code.gs` with the supplied `Code.gs`.
3. Replace `index.html` with `index_fixed.html` (or rename it to `index.html`).
4. Run `setupSystem()` once from Apps Script editor and authorize Drive/Sheets permissions.
5. Deploy a new Web App version.
6. Confirm the frontend `WEB_APP_URL` points to the current Web App deployment URL.

## Important
Do not paste the generated HTML into Apps Script as a `.gs` file. `index_fixed.html` must remain an HTML file.
