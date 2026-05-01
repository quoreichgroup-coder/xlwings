# Dev Sync — Editing VBA in VS Code

VBA code lives inside the .xlsm workbook by default. The DevSync module lets you export modules to plain `.bas` text files, edit them in VS Code, and import them back — giving you syntax highlighting, Git version control, and a proper text editor for your VBA code.

## Workflow Overview

```
Excel VBA Editor                  VS Code (Git-tracked)
      │                                    │
      │── ExportVBAModules() ─────────────▶│  sap_analyst_framework/vba/*.bas
      │                                    │  (edit, search, version control)
      │◀─ ImportVBAModules() ─────────────│
      │                                    │
```

## Prerequisites

- **Trust access to the VBA project object model** must be enabled (see setup_without_addin.md Step 2).
- **VS Code** installed with `code` on the system PATH. Install [VS Code](https://code.visualstudio.com/) and run `code --version` in a terminal to verify.
- Recommended VS Code extension: **VBA** by serkonda7 (syntax highlighting for .bas files).

## Exporting VBA Modules to Files

1. Open the .xlsm workbook.
2. Open the VBE (Alt+F11).
3. Run `modDevSync.ExportVBAModules()` from the Immediate window:
   ```
   modDevSync.ExportVBAModules
   ```
   Or assign it to a button on the HOME sheet.

All standard modules are saved as `.bas` files in the `vba/` folder next to the workbook.

## Opening the Project in VS Code

Run from the Immediate window or a button:
```
modDevSync.OpenProjectInVSCode
```

This runs `code "<path_to_vba_folder>"` via Shell, opening the folder in VS Code.

## Editing .bas Files in VS Code

- `.bas` files are plain text — edit them freely.
- The first line `Attribute VB_Name = "modXxx"` must be preserved for correct import.
- Use Git to track changes: `git add vba/` → `git commit -m "..."`

## Importing .bas Files Back to Excel

After editing, import back from the Immediate window or a button:
```
modDevSync.ImportVBAModules
```

This removes the old module from the VBProject and re-imports from the file. **All in-memory changes in the VBE for that module will be overwritten.**

## Round-Trip Sync

To export, then immediately re-import (useful to verify consistency):
```
modDevSync.SyncVBAModules
```

## Exporting a Single Module

From the VBE, with the module open:
```
modDevSync.ExportCurrentModule
```

## Git Version Control

Since `.bas` files are plain text, you can version control them with Git:

```bash
cd "path/to/workbook_folder"
git init
git add vba/
git commit -m "Initial VBA module export"
```

Recommended `.gitignore`:
```
Exports/
Archive/
Logs/
*.tmp
~$*.xlsm
```

## VS Code Recommended Extensions

| Extension | Purpose |
|---|---|
| VBA (serkonda7) | Syntax highlighting for .bas |
| GitLens | Enhanced Git history |
| Excel Viewer | Preview .xlsx files inline |
| Python | For editing the python/ scripts |

## Important Notes

- Always export before editing in VS Code to get the latest in-memory VBA state.
- Always import after editing to apply changes to the workbook.
- Class modules (.cls) and UserForms (.frm) require separate handling — modDevSync currently exports only standard modules (.bas). Add `.cls` support by filtering `vbComp.Type = 2` if needed.
- The workbook must be saved after import for changes to persist across sessions.
