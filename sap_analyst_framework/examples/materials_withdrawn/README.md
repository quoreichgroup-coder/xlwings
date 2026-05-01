# Materials Withdrawn Example

Tracks material withdrawals from plant stock — goods issues and reversals.

## Transaction

`MB51` — Material Document List

## Key Features

- Uses Python (`UsePython = YES`) for data cleaning and SAP number format conversion.
- Filters movement types 261 (Goods Issue) and 262 (Reversal) separately.
- Calculates withdrawal counts, quantities, and reversal rates.

## KPIs

| KPI | Formula | Filter | Target |
|---|---|---|---|
| Total Withdrawn Qty | SUM Quantity | — | HOME!E8 |
| Unique Materials | COUNT_UNIQUE Material | — | HOME!E9 |
| Withdrawal Count | COUNT Material | — | HOME!E10 |
| Goods Issue Count | COUNT Movement Type | = 261 | HOME!E11 |
| Total GI Qty | SUM Quantity | Movement Type = 261 | HOME!E12 |
| Reversal Count | COUNT Movement Type | = 262 | HOME!E13 |
| Max Withdrawal Qty | MAX Quantity | — | HOME!E14 |
| Avg Withdrawal Qty | AVERAGE Quantity | — | HOME!E15 |

## Python Processing

With `UsePython = YES`, the Python pipeline runs `cleaner.py` to:
- Convert SAP comma-decimal quantities to floats.
- Parse posting date columns.
- Drop empty rows from the ALV export header.

The cleaned data is written back to DATA before VBA runs the KPI engine.

## DATA Sheet Expected Columns

- `Material`
- `Material Description`
- `Quantity`
- `Unit`
- `Movement Type`
- `Posting Date`
- `Plant`
- `Storage Location`
- `Batch`

## Setup

1. Copy SETTINGS.csv → SETTINGS sheet.
2. Copy REPORT_CONFIG.csv → REPORT_CONFIG sheet.
3. If using Python: verify `PythonPath` in SETTINGS or ensure `python` is on PATH.
4. Run the pipeline from HOME.
