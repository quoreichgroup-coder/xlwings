# UI Components & Loading States

The framework provides a reusable component library (`modComponents`) and a loading state manager (`modUIState`). Components can be placed manually or generated automatically from the UI_CONFIG sheet.

## UI Modes

| Mode | Behaviour |
|---|---|
| `MANUAL` | Analyst designs the HOME sheet by hand. VBA only writes values to configured target cells. Layout is never touched. |
| `AUTO` | VBA reads UI_CONFIG and creates / recreates components at pipeline start. |

Set `UI_MODE` in the SETTINGS sheet.

## Loader Types

| Loader | Shape |
|---|---|
| `BAR` | A horizontal bar shape (`ProgressBar_Fill`) that grows from left to right. |
| `CIRCLE` | A circular shape (`LoaderCircle`) that rotates proportionally to progress. |

Set `LOADER_TYPE` in the SETTINGS sheet.

---

## Components

### Button — `CreateButton`

A rounded-rectangle shape with Gold fill and a macro assignment.

```vba
modComponents.CreateButton _
    ws       := Sheet1, _
    topRow   := 5, _
    leftCol  := 2, _
    caption  := "Run Report", _
    colSpan  := 3, _
    rowSpan  := 1, _
    macroName := "modMain.RunMain"
```

**UI_CONFIG row:**
```
BUTTON | HOME | 5 | 2 | 3 | 1 | Run Report | | modMain.RunMain |
```

---

### KPI Card — `CreateKpiCard`

A dark Charcoal block with a Gold label on top and a large white value below.  
The value cell reference is a formula pointing at the calculated KPI cell.

```vba
modComponents.CreateKpiCard _
    ws        := Sheet1, _
    topRow    := 8, _
    leftCol   := 2, _
    label     := "Total Qty", _
    valueCell := "HOME!E10", _
    colSpan   := 2, _
    rowSpan   := 3
```

**UI_CONFIG row:**
```
KPI_CARD | HOME | 8 | 2 | 2 | 3 | Total Qty | HOME!E10 | |
```

---

### Status Box — `CreateStatusBox`

A coloured cell block showing a status label.

Styles: `OK` (green), `ERROR` (red), `WARN` (yellow), or any other string (light gray).

```vba
modComponents.CreateStatusBox ws, 12, 2, "Data loaded", 3, 1, "OK"
```

---

### Section Header — `CreateSectionHeader`

A full-width merged cell with Charcoal background and Gold uppercase title.

```vba
modComponents.CreateSectionHeader ws, 3, 1, "KPI Summary", 8
```

---

### Progress Bar — `CreateProgressBar`

Creates two stacked shapes: a light gray background bar and a Gold fill bar.  
`modUIState.UpdateLoading` resizes the fill bar during the pipeline.

Shape names:
- `ProgressBar_BG` — background
- `ProgressBar_Fill` — animated fill

```vba
modComponents.CreateProgressBar ws, 20, 2, 0, 8, 1
```

---

### Progress Circle — `CreateProgressCircle`

A circular shape (`LoaderCircle`) that rotates to indicate progress.

```vba
modComponents.CreateProgressCircle ws, 20, 2, 0, 40
```

---

### Alert Box — `CreateAlertBox`

Equivalent to a merged status box — used for prominent messages.

```vba
modComponents.CreateAlertBox ws, 15, 2, "Warning: data may be incomplete", 6, 2, "WARN"
```

---

### Navigation Button — `CreateNavigationButton`

A button that navigates to another sheet. The `targetSheet` parameter sets the OnAction macro.  
Add a public sub `modUI.NavigateTo_<SheetName>` for each navigable sheet.

```vba
modComponents.CreateNavigationButton ws, 5, 8, "→ Settings", 2, 1, "SETTINGS"
```

---

## UI_CONFIG Sheet Reference

Headers (row 1):

| Column | Field | Example |
|---|---|---|
| A | Component | `BUTTON`, `KPI_CARD`, `STATUS_BOX`, `SECTION_HEADER`, `NAV_BUTTON`, `PROGRESS_BAR`, `PROGRESS_CIRCLE`, `ALERT_BOX` |
| B | Sheet | `HOME` |
| C | Row | `5` |
| D | Col | `2` |
| E | Width (col span) | `3` |
| F | Height (row span) | `1` |
| G | Label | `Run Report` |
| H | Value Cell | `HOME!E10` (KPI_CARD only) |
| I | Style | `OK`, `ERROR`, `WARN` |
| J | Color Override | (leave blank for defaults) |

---

## modUIState — Loading Lifecycle

```vba
' At pipeline start
modUIState.StartLoading "Initialising..."

' During each step
modUIState.UpdateLoading 30, "Connecting to SAP..."
modUIState.UpdateLoading 60, "Importing data..."

' On completion
modUIState.StopLoading "Report complete."

' Feedback
modUIState.ShowSuccess "All KPIs calculated."
modUIState.ShowError "SAP connection failed."
```

The status message is written to cell `HOME!B3`.  
The percentage is written to cell `HOME!B4`.

To move these, change the constants `STATUS_CELL` and `PROGRESS_CELL` in `modUIState`.

---

## Colour Reference

| Token | Hex | RGB |
|---|---|---|
| Gold | #CCAA66 | (204, 170, 102) |
| Charcoal | #1F1F1F | (31, 31, 31) |
| LightGray | #FAFAFA | (250, 250, 250) |
| White | #FFFFFF | (255, 255, 255) |
| ErrorRed | — | (220, 53, 69) |
| SuccessGreen | — | (40, 167, 69) |
| WarnYellow | — | (255, 193, 7) |
| InfoBlue | — | (0, 123, 255) |

All tokens are available as properties of `modTheme`:
```vba
ws.Range("A1").Interior.Color = modTheme.Gold
```
