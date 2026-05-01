# Templates

## SAP_Analyst_Framework.xlsm

Create this workbook manually in Excel with the following sheets (in order):

1. **HOME** — Dashboard: add a "Run Report" shape/button and assign macro `modMain.RunMain`
2. **SETTINGS** — Key-value pairs; headers: `Key` (A1), `Value` (B1)
3. **REPORT_CONFIG** — KPI definitions; headers: `KPI_Name`, `Formula_Type`, `Column_Name`, `Target_Sheet`, `Target_Cell`, `Filter_Column`, `Filter_Value`, `Label`
4. **UI_CONFIG** — UI layout (AUTO mode); headers: `Component`, `Sheet`, `Row`, `Col`, `Width`, `Height`, `Label`, `Value_Cell`, `Style`, `Color_Override`
5. **DATA** — Leave blank; the framework imports SAP exports here
6. **LOG** — Leave blank; the framework writes run logs here
7. **HISTORY** — Leave blank; headers auto-created: `Timestamp`, `Report`, `Plant`, `Date From`, `Date To`, `User`

After creating the workbook:
1. Save as `.xlsm` (macro-enabled workbook).
2. Import all `.bas` files from the `vba/` folder using the VBE or `modDevSync.ImportVBAModules`.
3. Fill in the SETTINGS sheet from one of the examples.
4. Add your first REPORT_CONFIG rows.
5. Click Run Report.

## Why is the .xlsm file not included?

Binary Office files (.xlsm) cannot be meaningfully version-controlled in Git — every save generates a new binary diff even for trivial changes. Keeping only the text-based `.bas` and `.csv` configuration files in Git ensures clean history and easy code review.
