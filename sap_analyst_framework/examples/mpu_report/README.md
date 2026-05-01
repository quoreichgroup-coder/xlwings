# MPU Report Example

MPU (Material Planning Unit) report — tracks weekly material usage across plants.

## Transaction

`MCBA` — Inventory Controlling: Material Analysis

## What It Does

1. Opens MB52 variant `/MPU_WEEKLY` in SAP.
2. Applies layout `/MPU_LAYOUT`.
3. Sets date range to Previous Week (auto-resolved).
4. Exports the ALV grid to an Excel file.
5. Calculates KPIs: total usage, unique materials, movement count, plant breakdowns.
6. Updates HOME sheet cells.

## KPIs

| KPI | Formula | Target Cell |
|---|---|---|
| Total Usage Qty | SUM(Quantity) | HOME!E8 |
| Unique Materials | COUNT_UNIQUE(Material) | HOME!E9 |
| Movement Count | COUNT(Movement Type) | HOME!E10 |
| Avg Qty per Material | AVERAGE(Quantity) | HOME!E11 |
| Max Single Withdrawal | MAX(Quantity) | HOME!E12 |
| Min Single Withdrawal | MIN(Quantity) | HOME!E13 |
| Usage Qty Plant 1000 | SUM(Quantity) filtered Plant=1000 | HOME!E14 |
| Usage Qty Plant 2000 | SUM(Quantity) filtered Plant=2000 | HOME!E15 |

## Setup

1. Copy SETTINGS.csv content into the SETTINGS sheet (columns A and B).
2. Copy REPORT_CONFIG.csv content into the REPORT_CONFIG sheet.
3. Verify that the DATA sheet column headers match the Column_Name values above after an export.
4. Click Run Report on HOME.

## DATA Sheet Expected Columns

After export, the DATA sheet should contain columns matching (case-insensitive):
- `Material`
- `Quantity`
- `Plant`
- `Movement Type`
- `Posting Date`
- `Storage Location`

Column names must match the SAP ALV layout. Adjust REPORT_CONFIG if your layout uses different names.
