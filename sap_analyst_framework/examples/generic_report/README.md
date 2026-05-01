# Generic SAP Report Example

A starter template for any SAP ALV report. Copy this example and customise it for your specific TCode.

## How to Use This Template

### 1. Choose Your TCode

Update `TCode` in SETTINGS.csv to match your SAP report:

| TCode | Report |
|---|---|
| MB52 | Warehouse Stocks |
| MB51 | Material Documents |
| ME2M | Purchase Orders by Material |
| ME2L | Purchase Orders by Vendor |
| MM60 | Inventory Turnover |
| MCBA | Inventory Controlling |
| CO03 | Production Order Display |

### 2. Create Your Variant and Layout in SAP

1. Run the TCode manually in SAP.
2. Fill in your usual selection parameters.
3. Go to **Goto → Variants → Save as Variant** — give it a name like `/MY_VARIANT`.
4. Set up the ALV column layout you need.
5. Save the layout: **Settings → Layout → Save** — name it `/MY_LAYOUT`.

### 3. Update REPORT_CONFIG

Replace the placeholder column names with the actual headers that appear in your exported DATA sheet after a first manual run. Run once manually (click Run Report), then check column A row 1 of the DATA sheet for exact header names.

### 4. Choose DateMode

| Mode | Use when |
|---|---|
| CUSTOM | You want a specific fixed date range |
| PREVIOUS_WEEK | Weekly recurring report |
| CURRENT_WEEK | Monitor week-to-date |
| MONTH_TO_DATE | Monthly rolling report |
| YEAR_TO_DATE | YTD cumulative |

### 5. Run and Iterate

- Run the pipeline and check the LOG sheet for any errors.
- Adjust REPORT_CONFIG column names to match the DATA headers.
- Adjust target cells to match your HOME sheet design.

## Tips

- Use `COUNT_UNIQUE` to identify the number of distinct materials, vendors, or orders.
- Use `Filter_Column` + `Filter_Value` to split one KPI by category without creating separate reports.
- Combine `FORMULA` type with a pandas expression (Python mode) for complex multi-column calculations.
- Keep `UsePython = NO` for simple reports to avoid the Python dependency.

## Troubleshooting Column Names

If a KPI shows `COL_NOT_FOUND`:
1. Check the DATA sheet header row (row 1) after a run.
2. Copy the exact column name into REPORT_CONFIG Column_Name.
3. The match is case-insensitive but spaces and special characters must be exact.
