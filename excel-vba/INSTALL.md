# Excel Row Highlighter — Installation Guide

## Quick Install (Personal.xlsb)

### Step 1: Open Personal.xlsb

1. Open Excel
2. Press `Alt+F11` to open VBA Editor
3. In the Project Explorer, find `VBAProject (PERSONAL.XLSB)`
   - If it doesn't exist, you need to create it first:
     1. In Excel, record a dummy macro (Developer → Record Macro → store in "Personal Macro Workbook" → OK → Stop)
     2. This creates `Personal.xlsb` automatically
     3. Delete the dummy macro afterward

### Step 2: Import Standard Modules

1. In VBA Editor, right-click `PERSONAL.XLSB` → Import File...
2. Import `modSettings.bas`
3. Import `modHighlighter.bas`

### Step 3: Import Class Module

1. Right-click `PERSONAL.XLSB` → Import File...
2. Import `clsAppEvents.cls`

### Step 4: Create the UserForm

**Note:** The `.frm` file cannot be directly imported without its binary `.frx` companion. Create it manually:

1. Right-click `PERSONAL.XLSB` → Insert → UserForm
2. Set Name to `frmSettings` in Properties
3. Set Caption to `Excel Row Highlighter Settings`
4. Add these controls (use Toolbox):

| Control Type | Name | Caption/Text |
|---|---|---|
| CheckBox | `chkRowLine` | Row Line |
| CheckBox | `chkColLine` | Col Line |
| CheckBox | `chkRowFill` | Row Fill |
| CheckBox | `chkColFill` | Col Fill |
| Label | `lblRowLineSize` | Size: |
| Label | `lblColLineSize` | Size: |
| Label | `lblRowFillOp` | Opacity: |
| Label | `lblColFillOp` | Opacity: |
| TextBox | `txtRowLineSize` | 3.25 |
| TextBox | `txtColLineSize` | 3.25 |
| TextBox | `txtRowFillOpacity` | 0.05 |
| TextBox | `txtColFillOpacity` | 0.05 |
| CommandButton | `btnRowLineColor` | #c2185b |
| CommandButton | `btnColLineColor` | #c2185b |
| CommandButton | `btnRowFillColor` | #c2185b |
| CommandButton | `btnColFillColor` | #c2185b |
| CommandButton | `btnReset` | Reset Defaults |
| CommandButton | `btnApply` | Apply & Close |

5. Layout (4 rows + 1 button row):
   ```
   Row 1: [chkRowLine]  [Size: txtRowLineSize]  [btnRowLineColor]
   Row 2: [chkColLine]  [Size: txtColLineSize]  [btnColLineColor]
   Row 3: [chkRowFill]  [Opacity: txtRowFillOpacity]  [btnRowFillColor]
   Row 4: [chkColFill]  [Opacity: txtColFillOpacity]  [btnColFillColor]
   Row 5: [btnReset]                             [btnApply]
   ```

6. Double-click the form → paste the code from `frmSettings.frm` (the VBA code section after the `Option Explicit` line)

### Step 5: Setup ThisWorkbook

1. In `PERSONAL.XLSB`, double-click `ThisWorkbook`
2. Paste the code from `ThisWorkbook.cls` (only the VBA code, not the VERSION/Attribute headers)

### Step 6: Save & Restart

1. Press `Ctrl+S` to save Personal.xlsb
2. Close and reopen Excel
3. The highlighter should now be active!

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+Shift+R` | Toggle row highlight |
| `Ctrl+Shift+C` | Toggle column highlight |
| `Ctrl+Shift+A` | Toggle all highlights on/off |
| `Ctrl+Shift+H` | Open settings dialog |

---

## Features

- **Row Line** — colored borders at top/bottom of active row
- **Col Line** — colored borders at left/right of active column
- **Row Fill** — semi-transparent fill across the entire row
- **Col Fill** — semi-transparent fill down the entire column
- **Per-setting colors** — independent colors for each feature
- **Merged cell support** — highlights cover the full merged area
- **Protected sheet safe** — silently skips protected sheets
- **Global** — works on all open workbooks automatically

---

## Known Limitations

1. **Scroll:** Highlights only update on cell selection, not on scroll. After scrolling, click any cell to refresh.
2. **Undo stack:** Shape operations affect Excel's undo history. `Ctrl+Z` may undo a shape operation instead of your last edit.
3. **Performance:** Sheets with many shapes may see slight delay during selection changes.

---

## Uninstall

1. Open VBA Editor (`Alt+F11`)
2. Under `PERSONAL.XLSB`, delete these modules:
   - `modSettings`
   - `modHighlighter`
   - `clsAppEvents`
   - `frmSettings`
3. Clear `ThisWorkbook` code in PERSONAL.XLSB
4. Clear registry settings — run in Immediate Window (`Ctrl+G`):
   ```
   DeleteSetting "ExcelRowHighlighter"
   ```
5. Save and restart Excel

---

## Troubleshooting

**Highlights don't appear:**
- Check that macros are enabled (File → Options → Trust Center → Macro Settings)
- Verify Personal.xlsb exists in `%APPDATA%\Microsoft\Excel\XLSTART\`
- Press `Alt+F11`, check Immediate Window (`Ctrl+G`) for error messages starting with "RH"

**Shortcuts don't work:**
- Restart Excel (shortcuts register on startup)
- Check for conflicts with other add-ins

**Color picker shows error:**
- Make sure at least one workbook is open (not just Personal.xlsb)
