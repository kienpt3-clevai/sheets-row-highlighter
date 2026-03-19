# Excel Row Highlighter (VBA Personal.xlsb) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Highlight active cell's row and column in Excel desktop (row line, col line, row fill, col fill) — runs globally on all workbooks via Personal.xlsb.

**Architecture:** Application-level event handler (`WithEvents`) in a class module catches `SheetSelectionChange` across all open workbooks. A highlighter module draws/clears named Shapes on each selection change. Settings stored in Windows Registry via `SaveSetting`/`GetSetting`. A UserForm provides the settings UI.

**Tech Stack:** VBA (Excel), Shapes API, Registry (`SaveSetting`/`GetSetting`), UserForm

---

## File Structure

All modules live inside `Personal.xlsb`. Source code is also exported to `excel-vba/` for version control.

| VBA Module | Type | File (exported) | Responsibility |
|---|---|---|---|
| `ThisWorkbook` | Document Module | `excel-vba/ThisWorkbook.cls` | Auto-initialize app event handler on `Workbook_Open` |
| `clsAppEvents` | Class Module | `excel-vba/clsAppEvents.cls` | `WithEvents Application` — catches `SheetSelectionChange` |
| `modHighlighter` | Standard Module | `excel-vba/modHighlighter.bas` | Draw/clear highlight shapes (lines + fills) |
| `modSettings` | Standard Module | `excel-vba/modSettings.bas` | Read/write settings from Registry, provide defaults |
| `frmSettings` | UserForm | `excel-vba/frmSettings.frm` | Settings UI (toggles, colors, sizes, opacity) |

### Shape Naming Convention

All shapes created by the highlighter use prefix `"RH_"`:
- `RH_RowLineTop`, `RH_RowLineBottom` — row border lines
- `RH_ColLineLeft`, `RH_ColLineRight` — column border lines
- `RH_RowFill` — row background fill
- `RH_ColFill` — column background fill

This allows fast identification and cleanup without affecting user shapes.

### How Highlighting Works

```
Excel Startup
  → Personal.xlsb Workbook_Open
    → Creates clsAppEvents instance
      → WithEvents App As Application

User selects a cell (any workbook)
  → App_SheetSelectionChange fires
    → modHighlighter.ClearHighlights (delete all RH_* shapes)
    → modHighlighter.DrawHighlights:
      1. Get ActiveCell.Row, ActiveCell.Column
      2. Get visible range boundaries (ActiveWindow.VisibleRange)
      3. If rowLineEnabled: draw 2 horizontal lines (top/bottom of row)
      4. If colLineEnabled: draw 2 vertical lines (left/right of column)
      5. If rowFillEnabled: draw rectangle across row with transparency
      6. If colFillEnabled: draw rectangle down column with transparency
```

### Settings Stored in Registry

Registry App Name: `"ExcelRowHighlighter"`

| Key | Type | Default | Description |
|---|---|---|---|
| `RowLineEnabled` | Boolean | `True` | Toggle row border lines |
| `ColLineEnabled` | Boolean | `True` | Toggle column border lines |
| `RowFillEnabled` | Boolean | `True` | Toggle row background fill |
| `ColFillEnabled` | Boolean | `True` | Toggle column background fill |
| `RowLineColor` | String | `"#c2185b"` | Row line color (hex) |
| `ColLineColor` | String | `"#c2185b"` | Column line color (hex) |
| `RowFillColor` | String | `"#c2185b"` | Row fill color (hex) |
| `ColFillColor` | String | `"#c2185b"` | Column fill color (hex) |
| `RowLineSize` | Double | `3.25` | Row line weight (pt) |
| `ColLineSize` | Double | `3.25` | Column line weight (pt) |
| `RowFillOpacity` | Double | `0.05` | Row fill opacity (0-0.5) |
| `ColFillOpacity` | Double | `0.05` | Column fill opacity (0-0.5) |

---

## Task 1: Project Setup & Settings Module

**Files:**
- Create: `excel-vba/modSettings.bas`

- [ ] **Step 1: Create the exported VBA source directory**

```bash
mkdir -p excel-vba
```

- [ ] **Step 2: Write modSettings.bas with defaults and Registry read/write**

```vba
Attribute VB_Name = "modSettings"
Option Explicit

' === Registry constants ===
Public Const APP_NAME As String = "ExcelRowHighlighter"
Public Const SEC_GENERAL As String = "General"

' === Default values (matching Chrome extension) ===
Public Const DEF_ROW_LINE_ENABLED As Boolean = True
Public Const DEF_COL_LINE_ENABLED As Boolean = True
Public Const DEF_ROW_FILL_ENABLED As Boolean = True
Public Const DEF_COL_FILL_ENABLED As Boolean = True
Public Const DEF_ROW_LINE_COLOR As String = "#c2185b"
Public Const DEF_COL_LINE_COLOR As String = "#c2185b"
Public Const DEF_ROW_FILL_COLOR As String = "#c2185b"
Public Const DEF_COL_FILL_COLOR As String = "#c2185b"
Public Const DEF_ROW_LINE_SIZE As Double = 3.25
Public Const DEF_COL_LINE_SIZE As Double = 3.25
Public Const DEF_ROW_FILL_OPACITY As Double = 0.05
Public Const DEF_COL_FILL_OPACITY As Double = 0.05

' === Cached settings (loaded once, updated by UserForm) ===
Public RowLineEnabled As Boolean
Public ColLineEnabled As Boolean
Public RowFillEnabled As Boolean
Public ColFillEnabled As Boolean
Public RowLineColor As Long    ' OLE color
Public ColLineColor As Long
Public RowFillColor As Long
Public ColFillColor As Long
Public RowLineSize As Double
Public ColLineSize As Double
Public RowFillOpacity As Double
Public ColFillOpacity As Double

' --- Hex to OLE Color conversion ---
Public Function HexToOLE(ByVal hex As String) As Long
    ' Convert "#RRGGBB" to OLE color (Long = &HBBGGRR)
    If Left$(hex, 1) = "#" Then hex = Mid$(hex, 2)
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid$(hex, 1, 2))
    g = CLng("&H" & Mid$(hex, 3, 2))
    b = CLng("&H" & Mid$(hex, 5, 2))
    HexToOLE = RGB(r, g, b)
End Function

' --- OLE Color to Hex conversion ---
Public Function OLEToHex(ByVal oleColor As Long) As String
    Dim r As Long, g As Long, b As Long
    r = oleColor Mod 256
    g = (oleColor \ 256) Mod 256
    b = (oleColor \ 65536) Mod 256
    OLEToHex = "#" & Right$("0" & Hex$(r), 2) & _
                      Right$("0" & Hex$(g), 2) & _
                      Right$("0" & Hex$(b), 2)
End Function

' --- Read a string from Registry with default ---
Private Function RegGet(ByVal key As String, ByVal def As String) As String
    On Error Resume Next
    RegGet = GetSetting(APP_NAME, SEC_GENERAL, key, def)
End Function

' --- Write a string to Registry ---
Private Sub RegSet(ByVal key As String, ByVal val As String)
    SaveSetting APP_NAME, SEC_GENERAL, key, val
End Sub

' --- Load all settings from Registry into public vars ---
Public Sub LoadSettings()
    RowLineEnabled = CBool(RegGet("RowLineEnabled", CStr(DEF_ROW_LINE_ENABLED)))
    ColLineEnabled = CBool(RegGet("ColLineEnabled", CStr(DEF_COL_LINE_ENABLED)))
    RowFillEnabled = CBool(RegGet("RowFillEnabled", CStr(DEF_ROW_FILL_ENABLED)))
    ColFillEnabled = CBool(RegGet("ColFillEnabled", CStr(DEF_COL_FILL_ENABLED)))
    RowLineColor = HexToOLE(RegGet("RowLineColor", DEF_ROW_LINE_COLOR))
    ColLineColor = HexToOLE(RegGet("ColLineColor", DEF_COL_LINE_COLOR))
    RowFillColor = HexToOLE(RegGet("RowFillColor", DEF_ROW_FILL_COLOR))
    ColFillColor = HexToOLE(RegGet("ColFillColor", DEF_COL_FILL_COLOR))
    RowLineSize = CDbl(RegGet("RowLineSize", CStr(DEF_ROW_LINE_SIZE)))
    ColLineSize = CDbl(RegGet("ColLineSize", CStr(DEF_COL_LINE_SIZE)))
    RowFillOpacity = CDbl(RegGet("RowFillOpacity", CStr(DEF_ROW_FILL_OPACITY)))
    ColFillOpacity = CDbl(RegGet("ColFillOpacity", CStr(DEF_COL_FILL_OPACITY)))
End Sub

' --- Save all settings to Registry from public vars ---
Public Sub SaveSettings()
    RegSet "RowLineEnabled", CStr(RowLineEnabled)
    RegSet "ColLineEnabled", CStr(ColLineEnabled)
    RegSet "RowFillEnabled", CStr(RowFillEnabled)
    RegSet "ColFillEnabled", CStr(ColFillEnabled)
    RegSet "RowLineColor", OLEToHex(RowLineColor)
    RegSet "ColLineColor", OLEToHex(ColLineColor)
    RegSet "RowFillColor", OLEToHex(RowFillColor)
    RegSet "ColFillColor", OLEToHex(ColFillColor)
    RegSet "RowLineSize", CStr(RowLineSize)
    RegSet "ColLineSize", CStr(ColLineSize)
    RegSet "RowFillOpacity", CStr(RowFillOpacity)
    RegSet "ColFillOpacity", CStr(ColFillOpacity)
End Sub

' --- Reset all settings to defaults ---
Public Sub ResetSettings()
    RowLineEnabled = DEF_ROW_LINE_ENABLED
    ColLineEnabled = DEF_COL_LINE_ENABLED
    RowFillEnabled = DEF_ROW_FILL_ENABLED
    ColFillEnabled = DEF_COL_FILL_ENABLED
    RowLineColor = HexToOLE(DEF_ROW_LINE_COLOR)
    ColLineColor = HexToOLE(DEF_COL_LINE_COLOR)
    RowFillColor = HexToOLE(DEF_ROW_FILL_COLOR)
    ColFillColor = HexToOLE(DEF_COL_FILL_COLOR)
    RowLineSize = DEF_ROW_LINE_SIZE
    ColLineSize = DEF_COL_LINE_SIZE
    RowFillOpacity = DEF_ROW_FILL_OPACITY
    ColFillOpacity = DEF_COL_FILL_OPACITY
    SaveSettings
End Sub
```

- [ ] **Step 3: Verify — open VBA editor, paste module, run `LoadSettings` in Immediate Window**

```
Call modSettings.LoadSettings
? modSettings.RowLineEnabled     ' Expected: True
? modSettings.RowLineColor       ' Expected: 12721755 (OLE for #c2185b)
```

- [ ] **Step 4: Commit**

```bash
git add excel-vba/modSettings.bas
git commit -m "feat(excel): add settings module with Registry persistence"
```

---

## Task 2: Highlighter Module — Clear & Draw Shapes

**Files:**
- Create: `excel-vba/modHighlighter.bas`

- [ ] **Step 1: Write modHighlighter.bas — ClearHighlights**

```vba
Attribute VB_Name = "modHighlighter"
Option Explicit

Public Const SHAPE_PREFIX As String = "RH_"

' Track previous highlight to avoid redundant redraws
Private mLastRow As Long
Private mLastCol As Long
Private mLastSheet As String

' --- Delete all RH_* shapes from the given sheet ---
Public Sub ClearHighlights(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim i As Long
    ' Iterate backwards to safely delete
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            shp.Delete
        End If
    Next i
End Sub
```

- [ ] **Step 2: Verify — create a test shape named "RH_Test" manually, run ClearHighlights**

In Immediate Window:
```
ActiveSheet.Shapes.AddLine(10, 10, 200, 10).Name = "RH_Test"
Call modHighlighter.ClearHighlights(ActiveSheet)
? ActiveSheet.Shapes.Count   ' RH_Test should be gone
```

- [ ] **Step 3: Write DrawHighlights procedure**

```vba
' --- Main entry: draw highlights for the active cell ---
Public Sub DrawHighlights(ByVal ws As Worksheet, ByVal target As Range)
    Dim visRange As Range
    Dim cellRow As Long, cellCol As Long
    Dim rowTop As Double, rowBottom As Double, rowHeight As Double
    Dim colLeft As Double, colRight As Double, colWidth As Double
    Dim visLeft As Double, visTop As Double, visRight As Double, visBottom As Double
    Dim shp As Shape

    ' Skip if nothing enabled
    If Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled Or _
            modSettings.RowFillEnabled Or modSettings.ColFillEnabled) Then
        Exit Sub
    End If

    cellRow = target.Row
    cellCol = target.Column

    ' Get active cell geometry
    With ws.Cells(cellRow, cellCol)
        rowTop = .Top
        rowHeight = .Height
        rowBottom = rowTop + rowHeight
        colLeft = .Left
        colWidth = .Width
        colRight = colLeft + colWidth
    End With

    ' Get visible range geometry
    Set visRange = Application.ActiveWindow.VisibleRange
    With visRange
        visLeft = .Left
        visTop = .Top
        visRight = .Left + .Width
        visBottom = .Top + .Height
    End With

    Application.ScreenUpdating = False

    ' --- Row Fill ---
    If modSettings.RowFillEnabled And modSettings.RowFillOpacity > 0 Then
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
            visLeft, rowTop, visRight - visLeft, rowHeight)
        With shp
            .Name = SHAPE_PREFIX & "RowFill"
            .Fill.ForeColor.RGB = modSettings.RowFillColor
            .Fill.Transparency = 1# - modSettings.RowFillOpacity
            .Line.Visible = msoFalse
            .Placement = xlFreeFloating
        End With
    End If

    ' --- Col Fill ---
    If modSettings.ColFillEnabled And modSettings.ColFillOpacity > 0 Then
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
            colLeft, visTop, colWidth, visBottom - visTop)
        With shp
            .Name = SHAPE_PREFIX & "ColFill"
            .Fill.ForeColor.RGB = modSettings.ColFillColor
            .Fill.Transparency = 1# - modSettings.ColFillOpacity
            .Line.Visible = msoFalse
            .Placement = xlFreeFloating
        End With
    End If

    ' --- Row Lines (top + bottom) ---
    If modSettings.RowLineEnabled Then
        ' Top line
        Set shp = ws.Shapes.AddLine( _
            visLeft, rowTop, visRight, rowTop)
        FormatLineShape shp, SHAPE_PREFIX & "RowLineTop", _
            modSettings.RowLineColor, modSettings.RowLineSize

        ' Bottom line
        Set shp = ws.Shapes.AddLine( _
            visLeft, rowBottom, visRight, rowBottom)
        FormatLineShape shp, SHAPE_PREFIX & "RowLineBottom", _
            modSettings.RowLineColor, modSettings.RowLineSize
    End If

    ' --- Col Lines (left + right) ---
    If modSettings.ColLineEnabled Then
        ' Left line
        Set shp = ws.Shapes.AddLine( _
            colLeft, visTop, colLeft, visBottom)
        FormatLineShape shp, SHAPE_PREFIX & "ColLineLeft", _
            modSettings.ColLineColor, modSettings.ColLineSize

        ' Right line
        Set shp = ws.Shapes.AddLine( _
            colRight, visTop, colRight, visBottom)
        FormatLineShape shp, SHAPE_PREFIX & "ColLineRight", _
            modSettings.ColLineColor, modSettings.ColLineSize
    End If

    Application.ScreenUpdating = True
End Sub

' --- Format a line shape ---
Private Sub FormatLineShape(ByVal shp As Shape, ByVal shapeName As String, _
                            ByVal lineColor As Long, ByVal lineWeight As Double)
    With shp
        .Name = shapeName
        .Line.ForeColor.RGB = lineColor
        .Line.Weight = lineWeight
        .Line.Visible = msoTrue
        .Placement = xlFreeFloating
    End With
End Sub
```

- [ ] **Step 4: Verify — run DrawHighlights manually**

In Immediate Window:
```
Call modSettings.LoadSettings
Call modHighlighter.ClearHighlights(ActiveSheet)
Call modHighlighter.DrawHighlights(ActiveSheet, ActiveCell)
' Expected: see colored lines/fills around active cell's row and column
```

- [ ] **Step 5: Add HasSelectionChanged optimization**

```vba
' --- Check if selection actually changed (avoid redundant redraws) ---
Public Function HasSelectionChanged(ByVal ws As Worksheet, ByVal target As Range) As Boolean
    Dim sheetName As String
    sheetName = ws.Parent.Name & "!" & ws.Name
    If target.Row = mLastRow And target.Column = mLastCol And sheetName = mLastSheet Then
        HasSelectionChanged = False
    Else
        mLastRow = target.Row
        mLastCol = target.Column
        mLastSheet = sheetName
        HasSelectionChanged = True
    End If
End Function
```

- [ ] **Step 6: Commit**

```bash
git add excel-vba/modHighlighter.bas
git commit -m "feat(excel): add highlighter module with shape drawing/clearing"
```

---

## Task 3: Application Event Handler & Auto-Init

**Files:**
- Create: `excel-vba/clsAppEvents.cls`
- Create: `excel-vba/ThisWorkbook.cls`

- [ ] **Step 1: Write clsAppEvents.cls**

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrHandler

    ' Only handle Worksheets (not Chart sheets, etc.)
    If Not TypeOf Sh Is Worksheet Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh

    ' Skip if selection didn't actually change
    If Not modHighlighter.HasSelectionChanged(ws, target) Then Exit Sub

    ' Clear old highlights, draw new ones
    modHighlighter.ClearHighlights ws
    modHighlighter.DrawHighlights ws, target

    Exit Sub
ErrHandler:
    Debug.Print "RH Error: " & Err.Description
End Sub
```

- [ ] **Step 2: Write ThisWorkbook.cls (Personal.xlsb's Workbook_Open)**

```vba
' This goes into Personal.xlsb > ThisWorkbook
Option Explicit

Private mAppEvents As clsAppEvents

Private Sub Workbook_Open()
    modSettings.LoadSettings
    Set mAppEvents = New clsAppEvents
    Set mAppEvents.App = Application
End Sub
```

- [ ] **Step 3: Verify — close and reopen Excel, select cells in any workbook**

Expected behavior:
1. Excel opens → Personal.xlsb loads → `Workbook_Open` fires
2. Click any cell → row/column highlights appear
3. Click another cell → old highlights clear, new ones appear

- [ ] **Step 4: Commit**

```bash
git add excel-vba/clsAppEvents.cls excel-vba/ThisWorkbook.cls
git commit -m "feat(excel): add app event handler and auto-init on startup"
```

---

## Task 4: UserForm Settings UI

**Files:**
- Create: `excel-vba/frmSettings.frm`

- [ ] **Step 1: Design UserForm layout**

UserForm `frmSettings` — layout matching the Chrome extension popup:

```
┌─────────────────────────────────────────────────┐
│  Excel Row Highlighter Settings                 │
├─────────────────────────────────────────────────┤
│                                                 │
│  ☑ Row Line    Size: [2.0 ▼]   Color: [■ btn]  │
│  ☑ Col Line    Size: [2.0 ▼]   Color: [■ btn]  │
│  ☑ Row Fill    Opacity: [0.05]  Color: [■ btn]  │
│  ☑ Col Fill    Opacity: [0.05]  Color: [■ btn]  │
│                                                 │
│  [ Reset Defaults ]           [ Apply & Close ] │
│                                                 │
└─────────────────────────────────────────────────┘
```

Controls:
- `chkRowLine`, `chkColLine`, `chkRowFill`, `chkColFill` — CheckBox
- `txtRowLineSize`, `txtColLineSize` — TextBox (0.5-5)
- `txtRowFillOpacity`, `txtColFillOpacity` — TextBox (0-0.5)
- `btnRowLineColor`, `btnColLineColor`, `btnRowFillColor`, `btnColFillColor` — CommandButton (shows color dialog)
- `btnReset` — CommandButton
- `btnApply` — CommandButton

- [ ] **Step 2: Write frmSettings UserForm code**

```vba
' frmSettings code-behind
Option Explicit

Private Sub UserForm_Initialize()
    ' Load current settings into controls
    chkRowLine.Value = modSettings.RowLineEnabled
    chkColLine.Value = modSettings.ColLineEnabled
    chkRowFill.Value = modSettings.RowFillEnabled
    chkColFill.Value = modSettings.ColFillEnabled
    txtRowLineSize.Text = CStr(modSettings.RowLineSize)
    txtColLineSize.Text = CStr(modSettings.ColLineSize)
    txtRowFillOpacity.Text = CStr(modSettings.RowFillOpacity)
    txtColFillOpacity.Text = CStr(modSettings.ColFillOpacity)
    ' Set button colors
    UpdateColorButton btnRowLineColor, modSettings.RowLineColor
    UpdateColorButton btnColLineColor, modSettings.ColLineColor
    UpdateColorButton btnRowFillColor, modSettings.RowFillColor
    UpdateColorButton btnColFillColor, modSettings.ColFillColor
End Sub

Private Sub UpdateColorButton(ByVal btn As MSForms.CommandButton, ByVal oleColor As Long)
    btn.BackColor = oleColor
    btn.Caption = modSettings.OLEToHex(oleColor)
End Sub

Private Sub btnRowLineColor_Click()
    PickColor btnRowLineColor
End Sub
Private Sub btnColLineColor_Click()
    PickColor btnColLineColor
End Sub
Private Sub btnRowFillColor_Click()
    PickColor btnRowFillColor
End Sub
Private Sub btnColFillColor_Click()
    PickColor btnColFillColor
End Sub

Private Sub PickColor(ByVal btn As MSForms.CommandButton)
    ' Guard: need an open workbook for color dialog
    If ActiveWorkbook Is Nothing Then
        MsgBox "Please open a workbook first to use the color picker.", vbExclamation
        Exit Sub
    End If

    ' Use Excel's built-in color dialog
    With Application.Dialogs(xlDialogEditColor)
        ' Pre-load current color into palette slot 1
        Dim curColor As Long
        curColor = btn.BackColor
        ActiveWorkbook.Colors(1) = curColor
        If .Show(1) Then
            Dim newColor As Long
            newColor = ActiveWorkbook.Colors(1)
            UpdateColorButton btn, newColor
        End If
    End With
End Sub

Private Sub btnApply_Click()
    ' Read values from controls
    modSettings.RowLineEnabled = chkRowLine.Value
    modSettings.ColLineEnabled = chkColLine.Value
    modSettings.RowFillEnabled = chkRowFill.Value
    modSettings.ColFillEnabled = chkColFill.Value
    modSettings.RowLineSize = CDbl(txtRowLineSize.Text)
    modSettings.ColLineSize = CDbl(txtColLineSize.Text)
    modSettings.RowFillOpacity = CDbl(txtRowFillOpacity.Text)
    modSettings.ColFillOpacity = CDbl(txtColFillOpacity.Text)
    modSettings.RowLineColor = btnRowLineColor.BackColor
    modSettings.ColLineColor = btnColLineColor.BackColor
    modSettings.RowFillColor = btnRowFillColor.BackColor
    modSettings.ColFillColor = btnColFillColor.BackColor

    ' Save to Registry
    modSettings.SaveSettings

    ' Refresh current highlight
    If Not ActiveCell Is Nothing Then
        modHighlighter.ClearHighlights ActiveSheet
        modHighlighter.DrawHighlights ActiveSheet, ActiveCell
    End If

    Unload Me
End Sub

Private Sub btnReset_Click()
    modSettings.ResetSettings
    UserForm_Initialize  ' Refresh controls with defaults
End Sub
```

- [ ] **Step 3: Add a public entry point to show the form**

Add to `modSettings.bas`:

```vba
' --- Show the settings UserForm ---
Public Sub ShowSettings()
    frmSettings.Show vbModal
End Sub
```

- [ ] **Step 4: Verify — run `ShowSettings` from Immediate Window**

```
Call modSettings.ShowSettings
' Expected: UserForm appears with current settings
' Change a color → Apply → highlight updates immediately
```

- [ ] **Step 5: Commit**

```bash
git add excel-vba/frmSettings.frm
git commit -m "feat(excel): add settings UserForm with color picker and apply"
```

---

## Task 5: Keyboard Shortcuts & Quick Toggle

**Files:**
- Modify: `excel-vba/clsAppEvents.cls` (add keyboard handler)
- Modify: `excel-vba/modHighlighter.bas` (add toggle subs)

- [ ] **Step 1: Add toggle subroutines to modHighlighter.bas**

```vba
' --- Quick toggles (callable from shortcuts) ---
Public Sub ToggleRowLine()
    modSettings.RowLineEnabled = Not modSettings.RowLineEnabled
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Public Sub ToggleColLine()
    modSettings.ColLineEnabled = Not modSettings.ColLineEnabled
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Public Sub ToggleAll()
    Dim newState As Boolean
    newState = Not (modSettings.RowLineEnabled And modSettings.ColLineEnabled)
    modSettings.RowLineEnabled = newState
    modSettings.ColLineEnabled = newState
    modSettings.RowFillEnabled = newState
    modSettings.ColFillEnabled = newState
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Private Sub RefreshHighlight()
    If ActiveSheet Is Nothing Then Exit Sub
    ClearHighlights ActiveSheet
    If Not ActiveCell Is Nothing Then
        DrawHighlights ActiveSheet, ActiveCell
    End If
End Sub
```

- [ ] **Step 2: Register keyboard shortcuts in Workbook_Open**

Add to `ThisWorkbook.cls`:

```vba
Private Sub Workbook_Open()
    modSettings.LoadSettings
    Set mAppEvents = New clsAppEvents
    Set mAppEvents.App = Application

    ' Register keyboard shortcuts
    Application.OnKey "^+r", "modHighlighter.ToggleRowLine"    ' Ctrl+Shift+R
    Application.OnKey "^+c", "modHighlighter.ToggleColLine"    ' Ctrl+Shift+C
    Application.OnKey "^+a", "modHighlighter.ToggleAll"        ' Ctrl+Shift+A
    Application.OnKey "^+h", "modSettings.ShowSettings"        ' Ctrl+Shift+H
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Unregister shortcuts
    Application.OnKey "^+r"
    Application.OnKey "^+c"
    Application.OnKey "^+a"
    Application.OnKey "^+h"

    ' Clean up orphaned shapes from all open worksheets
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error Resume Next
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            modHighlighter.ClearHighlights ws
        Next ws
    Next wb
    On Error GoTo 0
End Sub
```

- [ ] **Step 3: Verify shortcuts**

1. Restart Excel
2. Press `Ctrl+Shift+R` → row lines toggle off/on
3. Press `Ctrl+Shift+A` → all highlights toggle off/on
4. Press `Ctrl+Shift+H` → settings form opens

- [ ] **Step 4: Commit**

```bash
git add excel-vba/clsAppEvents.cls excel-vba/modHighlighter.bas excel-vba/ThisWorkbook.cls
git commit -m "feat(excel): add keyboard shortcuts for quick toggles and settings"
```

---

## Task 6: Performance & Edge Cases

**Files:**
- Modify: `excel-vba/modHighlighter.bas`
- Modify: `excel-vba/clsAppEvents.cls`

- [ ] **Step 1: Add scroll handler — redraw on window scroll**

Add to `clsAppEvents.cls`:

```vba
Private Sub App_SheetActivate(ByVal Sh As Object)
    ' Redraw when switching sheets
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Dim ws As Worksheet
    Set ws = Sh
    modHighlighter.ClearHighlights ws
    If Not ActiveCell Is Nothing Then
        modHighlighter.DrawHighlights ws, ActiveCell
    End If
End Sub
```

- [ ] **Step 2: Handle protected sheets and errors**

Wrap `DrawHighlights` in error handling:

```vba
' At the top of DrawHighlights, after "Option Explicit":
Public Sub DrawHighlights(ByVal ws As Worksheet, ByVal target As Range)
    On Error GoTo ErrHandler

    ' Skip protected sheets (can't add shapes)
    If ws.ProtectDrawingObjects Then Exit Sub

    ' ... existing code ...

    Exit Sub
ErrHandler:
    ' Silently fail — don't break user's workflow
    Application.ScreenUpdating = True
End Sub
```

- [ ] **Step 3: Handle merged cells**

In `DrawHighlights`, replace the active cell geometry block:

```vba
    ' Get active cell geometry (handle merged cells)
    Dim mergedArea As Range
    Set mergedArea = target.MergeArea
    With mergedArea
        rowTop = .Top
        rowHeight = .Height
        rowBottom = rowTop + rowHeight
        colLeft = .Left
        colWidth = .Width
        colRight = colLeft + colWidth
    End With
```

- [ ] **Step 4: Verify edge cases**

1. Open a protected sheet → no error, no shapes drawn
2. Select a merged cell → highlight covers the full merged area
3. Scroll → highlights stay at correct position (shapes are anchored to cells)
4. Switch between sheets → highlights redraw correctly

- [ ] **Step 5: Commit**

```bash
git add excel-vba/modHighlighter.bas excel-vba/clsAppEvents.cls
git commit -m "fix(excel): handle scroll, protected sheets, and merged cells"
```

---

## Task 7: Installation Guide

**Files:**
- Create: `excel-vba/INSTALL.md`

- [ ] **Step 1: Write installation instructions**

```markdown
# Excel Row Highlighter — Installation Guide

## Quick Install (Personal.xlsb)

1. Open Excel
2. Press `Alt+F11` to open VBA Editor
3. In the Project Explorer, find `VBAProject (PERSONAL.XLSB)`
   - If it doesn't exist: File → New → save as `Personal.xlsb`
     in `%APPDATA%\Microsoft\Excel\XLSTART\`
4. Import modules:
   - Right-click → Import File → select `modSettings.bas`
   - Right-click → Import File → select `modHighlighter.bas`
   - Right-click → Import File → select `clsAppEvents.cls`
   - Right-click → Import File → select `frmSettings.frm`
5. Double-click `ThisWorkbook` under PERSONAL.XLSB
   - Paste the code from `ThisWorkbook.cls`
6. Close VBA Editor, save Personal.xlsb
7. Restart Excel

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+Shift+R` | Toggle row highlight |
| `Ctrl+Shift+C` | Toggle column highlight |
| `Ctrl+Shift+A` | Toggle all highlights |
| `Ctrl+Shift+H` | Open settings |

## Uninstall

1. Open VBA Editor (`Alt+F11`)
2. Delete all `RH_*` modules from PERSONAL.XLSB
3. Clear registry: Run `DeleteSetting "ExcelRowHighlighter"` in Immediate Window
```

- [ ] **Step 2: Commit**

```bash
git add excel-vba/INSTALL.md
git commit -m "docs(excel): add installation guide for Personal.xlsb setup"
```

---

## Known Limitations

1. **Scroll behavior:** VBA has no native scroll event. Highlights are drawn based on `VisibleRange` at the time of cell selection. After scrolling without selecting a new cell, highlights may appear cropped. They auto-correct on the next cell selection. (The Chrome extension handles this via continuous `requestAnimationFrame` polling, which VBA cannot replicate.)

2. **Performance on large sheets:** `ClearHighlights` iterates all shapes on the sheet. If the sheet has many user shapes, this loop runs slower. The `RH_` prefix filter keeps it targeted.

3. **Undo stack:** Adding/removing shapes pushes to Excel's undo stack. This means `Ctrl+Z` after selecting a cell may undo a shape deletion instead of the user's last edit. This is an inherent VBA limitation — there is no way to suppress undo entries for shape operations.

---

## Summary

| Task | What it delivers | Dependencies |
|------|-----------------|-------------|
| 1. Settings Module | Registry persistence, defaults, hex↔OLE conversion | None |
| 2. Highlighter Module | Shape drawing/clearing for lines + fills | Task 1 |
| 3. Event Handler | Auto-init, global SheetSelectionChange | Tasks 1, 2 |
| 4. UserForm | Settings UI with color picker | Tasks 1, 2 |
| 5. Keyboard Shortcuts | Quick toggles (Ctrl+Shift+R/C/A/S) | Tasks 1, 2, 3 |
| 6. Edge Cases | Protected sheets, merged cells, scroll | Tasks 1-3 |
| 7. Install Guide | Documentation for end users | All |
