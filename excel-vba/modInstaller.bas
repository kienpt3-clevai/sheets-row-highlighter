Attribute VB_Name = "modInstaller"
Option Explicit

' ============================================================
' EXCEL ROW HIGHLIGHTER - ONE-CLICK INSTALLER
'
' Prerequisites:
'   1. File > Options > Trust Center > Trust Center Settings
'      > Macro Settings > Tick "Trust access to the VBA project object model"
'   2. PERSONAL.XLSB must exist (record a dummy macro first if it doesn't)
'
' Usage:
'   1. Import this file into any workbook (Alt+F11 > right-click > Import)
'   2. In Immediate Window (Ctrl+G), run:  Call modInstaller.Install
'   3. Restart Excel. Done!
'
' To uninstall:  Call modInstaller.Uninstall
' ============================================================

Public Sub Install()
    Dim vbProj As Object

    On Error Resume Next
    Set vbProj = Workbooks("PERSONAL.XLSB").VBProject
    On Error GoTo 0

    If vbProj Is Nothing Then
        MsgBox "PERSONAL.XLSB not found!" & vbCrLf & vbCrLf & _
               "Create it first:" & vbCrLf & _
               "1. Developer > Record Macro" & vbCrLf & _
               "2. Store in: Personal Macro Workbook > OK" & vbCrLf & _
               "3. Stop Recording" & vbCrLf & _
               "4. Run this installer again", vbExclamation
        Exit Sub
    End If

    ' Check VBA project access
    On Error Resume Next
    Dim testAccess As Long
    testAccess = vbProj.VBComponents.Count
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Cannot access VBA project!" & vbCrLf & vbCrLf & _
               "Enable: File > Options > Trust Center > Trust Center Settings" & vbCrLf & _
               "> Macro Settings > Tick 'Trust access to the VBA project object model'", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Remove old installation if exists
    CleanupModules vbProj

    ' Install modules
    AddStandardModule vbProj, "modSettings", GetModSettingsCode()
    AddStandardModule vbProj, "modHighlighter", GetModHighlighterCode()
    AddClassModule vbProj, "clsAppEvents", GetClsAppEventsCode()
    SetThisWorkbookCode vbProj
    BuildSettingsForm vbProj

    MsgBox "Excel Row Highlighter installed successfully!" & vbCrLf & vbCrLf & _
           "Restart Excel to activate." & vbCrLf & vbCrLf & _
           "Shortcuts:" & vbCrLf & _
           "  Ctrl+Shift+R = Toggle row" & vbCrLf & _
           "  Ctrl+Shift+C = Toggle column" & vbCrLf & _
           "  Ctrl+Shift+A = Toggle all" & vbCrLf & _
           "  Ctrl+Shift+H = Settings", vbInformation
End Sub

Public Sub Uninstall()
    Dim vbProj As Object
    On Error Resume Next
    Set vbProj = Workbooks("PERSONAL.XLSB").VBProject
    On Error GoTo 0

    If vbProj Is Nothing Then
        MsgBox "PERSONAL.XLSB not found.", vbExclamation
        Exit Sub
    End If

    CleanupModules vbProj

    ' Clear ThisWorkbook code
    Dim twCode As Object
    Set twCode = vbProj.VBComponents("ThisWorkbook").CodeModule
    If twCode.CountOfLines > 0 Then twCode.DeleteLines 1, twCode.CountOfLines

    ' Clear registry
    On Error Resume Next
    DeleteSetting "ExcelRowHighlighter"
    On Error GoTo 0

    MsgBox "Excel Row Highlighter uninstalled. Restart Excel.", vbInformation
End Sub

' === Helper: Remove existing modules ===
Private Sub CleanupModules(ByVal vbProj As Object)
    Dim comp As Object
    Dim removeList As Object
    Set removeList = CreateObject("System.Collections.ArrayList")

    For Each comp In vbProj.VBComponents
        Select Case comp.Name
            Case "modSettings", "modHighlighter", "modFormBuilder", "clsAppEvents"
                removeList.Add comp.Name
            Case Else
                If comp.Type = 3 Then ' UserForm
                    If comp.Name = "frmSettings" Or Left$(comp.Name, 8) = "UserForm" Then
                        removeList.Add comp.Name
                    End If
                End If
        End Select
    Next comp

    Dim i As Long
    For i = 0 To removeList.Count - 1
        On Error Resume Next
        vbProj.VBComponents.Remove vbProj.VBComponents(removeList(i))
        On Error GoTo 0
    Next i
End Sub

' === Helper: Add a standard module ===
Private Sub AddStandardModule(ByVal vbProj As Object, ByVal modName As String, ByVal code As String)
    Dim comp As Object
    Set comp = vbProj.VBComponents.Add(1) ' vbext_ct_StdModule
    comp.Name = modName
    comp.CodeModule.AddFromString code
End Sub

' === Helper: Add a class module ===
Private Sub AddClassModule(ByVal vbProj As Object, ByVal clsName As String, ByVal code As String)
    Dim comp As Object
    Set comp = vbProj.VBComponents.Add(2) ' vbext_ct_ClassModule
    comp.Name = clsName
    comp.CodeModule.AddFromString code
End Sub

' === Helper: Set ThisWorkbook code ===
Private Sub SetThisWorkbookCode(ByVal vbProj As Object)
    Dim twCode As Object
    Set twCode = vbProj.VBComponents("ThisWorkbook").CodeModule
    If twCode.CountOfLines > 0 Then twCode.DeleteLines 1, twCode.CountOfLines
    twCode.AddFromString GetThisWorkbookCode()
End Sub

' === Helper: Build frmSettings UserForm ===
Private Sub BuildSettingsForm(ByVal vbProj As Object)
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm

    vbComp.Properties("Caption") = "Excel Row Highlighter Settings"
    vbComp.Properties("Width") = 340
    vbComp.Properties("Height") = 240

    Dim frm As Object
    Set frm = vbComp.Designer

    Dim yPos As Single, rowH As Single
    rowH = 30: yPos = 10

    ' Row 1: Row Line
    MakeCtrl frm, "Forms.CheckBox.1", "chkRowLine", "Row Line", 10, yPos, 75, 18
    MakeCtrl frm, "Forms.Label.1", "lblRLS", "Size:", 95, yPos + 2, 30, 14
    MakeCtrl frm, "Forms.TextBox.1", "txtRowLineSize", "2.25", 127, yPos, 45, 18
    MakeCtrl frm, "Forms.CommandButton.1", "btnRowLineColor", "#c2185b", 185, yPos, 135, 18
    yPos = yPos + rowH

    ' Row 2: Col Line
    MakeCtrl frm, "Forms.CheckBox.1", "chkColLine", "Col Line", 10, yPos, 75, 18
    MakeCtrl frm, "Forms.Label.1", "lblCLS", "Size:", 95, yPos + 2, 30, 14
    MakeCtrl frm, "Forms.TextBox.1", "txtColLineSize", "1.5", 127, yPos, 45, 18
    MakeCtrl frm, "Forms.CommandButton.1", "btnColLineColor", "#3399ff", 185, yPos, 135, 18
    yPos = yPos + rowH

    ' Row 3: Row Fill
    MakeCtrl frm, "Forms.CheckBox.1", "chkRowFill", "Row Fill", 10, yPos, 75, 18
    MakeCtrl frm, "Forms.Label.1", "lblRFO", "Opacity:", 95, yPos + 2, 40, 14
    MakeCtrl frm, "Forms.TextBox.1", "txtRowFillOpacity", "0.15", 137, yPos, 35, 18
    MakeCtrl frm, "Forms.CommandButton.1", "btnRowFillColor", "#c2185b", 185, yPos, 135, 18
    yPos = yPos + rowH

    ' Row 4: Col Fill
    MakeCtrl frm, "Forms.CheckBox.1", "chkColFill", "Col Fill", 10, yPos, 75, 18
    MakeCtrl frm, "Forms.Label.1", "lblCFO", "Opacity:", 95, yPos + 2, 40, 14
    MakeCtrl frm, "Forms.TextBox.1", "txtColFillOpacity", "0.05", 137, yPos, 35, 18
    MakeCtrl frm, "Forms.CommandButton.1", "btnColFillColor", "#3399ff", 185, yPos, 135, 18
    yPos = yPos + rowH + 10

    ' Row 5: Buttons
    MakeCtrl frm, "Forms.CommandButton.1", "btnReset", "Reset Defaults", 10, yPos, 90, 26
    MakeCtrl frm, "Forms.CommandButton.1", "btnSaveDefault", "Save Default", 125, yPos, 90, 26
    MakeCtrl frm, "Forms.CommandButton.1", "btnApply", "Apply && Close", 240, yPos, 90, 26

    ' Try to rename
    On Error Resume Next
    vbComp.Name = "frmSettings"
    On Error GoTo 0

    ' Add code
    Dim cm As Object
    Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString GetFrmSettingsCode()
End Sub

Private Sub MakeCtrl(ByVal frm As Object, ByVal progId As String, _
    ByVal sName As String, ByVal sCaption As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim ctrl As Object
    Set ctrl = frm.Controls.Add(progId, sName)
    ctrl.Left = x: ctrl.Top = y: ctrl.Width = w: ctrl.Height = h
    If progId = "Forms.TextBox.1" Then
        ctrl.Text = sCaption
    ElseIf progId = "Forms.CheckBox.1" Then
        ctrl.Caption = sCaption
        ctrl.Value = True
    Else
        ctrl.Caption = sCaption
    End If
End Sub

' ============================================================
' CODE PAYLOADS
' ============================================================

Private Function GetModSettingsCode() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf & vbCrLf
    c = c & "Public Const APP_NAME As String = ""ExcelRowHighlighter""" & vbCrLf
    c = c & "Public Const SEC_GENERAL As String = ""General""" & vbCrLf
    c = c & "Public Const SEC_DEFAULTS As String = ""CustomDefaults""" & vbCrLf & vbCrLf
    c = c & "Public Const DEF_ROW_LINE_ENABLED As Boolean = True" & vbCrLf
    c = c & "Public Const DEF_COL_LINE_ENABLED As Boolean = True" & vbCrLf
    c = c & "Public Const DEF_ROW_FILL_ENABLED As Boolean = True" & vbCrLf
    c = c & "Public Const DEF_COL_FILL_ENABLED As Boolean = True" & vbCrLf
    c = c & "Public Const DEF_ROW_LINE_COLOR As String = ""#c2185b""" & vbCrLf
    c = c & "Public Const DEF_COL_LINE_COLOR As String = ""#3399ff""" & vbCrLf
    c = c & "Public Const DEF_ROW_FILL_COLOR As String = ""#c2185b""" & vbCrLf
    c = c & "Public Const DEF_COL_FILL_COLOR As String = ""#3399ff""" & vbCrLf
    c = c & "Public Const DEF_ROW_LINE_SIZE As Double = 2.25" & vbCrLf
    c = c & "Public Const DEF_COL_LINE_SIZE As Double = 1.5" & vbCrLf
    c = c & "Public Const DEF_ROW_FILL_OPACITY As Double = 0.15" & vbCrLf
    c = c & "Public Const DEF_COL_FILL_OPACITY As Double = 0.05" & vbCrLf & vbCrLf
    c = c & "Public RowLineEnabled As Boolean" & vbCrLf
    c = c & "Public ColLineEnabled As Boolean" & vbCrLf
    c = c & "Public RowFillEnabled As Boolean" & vbCrLf
    c = c & "Public ColFillEnabled As Boolean" & vbCrLf
    c = c & "Public RowLineColor As Long" & vbCrLf
    c = c & "Public ColLineColor As Long" & vbCrLf
    c = c & "Public RowFillColor As Long" & vbCrLf
    c = c & "Public ColFillColor As Long" & vbCrLf
    c = c & "Public RowLineSize As Double" & vbCrLf
    c = c & "Public ColLineSize As Double" & vbCrLf
    c = c & "Public RowFillOpacity As Double" & vbCrLf
    c = c & "Public ColFillOpacity As Double" & vbCrLf & vbCrLf
    c = c & "Public gAppEvents As clsAppEvents" & vbCrLf & vbCrLf
    c = c & "Public Function HexToOLE(ByVal hex As String) As Long" & vbCrLf
    c = c & "    If Left$(hex, 1) = ""#"" Then hex = Mid$(hex, 2)" & vbCrLf
    c = c & "    Dim r As Long, g As Long, b As Long" & vbCrLf
    c = c & "    r = CLng(""&H"" & Mid$(hex, 1, 2))" & vbCrLf
    c = c & "    g = CLng(""&H"" & Mid$(hex, 3, 2))" & vbCrLf
    c = c & "    b = CLng(""&H"" & Mid$(hex, 5, 2))" & vbCrLf
    c = c & "    HexToOLE = RGB(r, g, b)" & vbCrLf
    c = c & "End Function" & vbCrLf & vbCrLf
    c = c & "Public Function OLEToHex(ByVal oleColor As Long) As String" & vbCrLf
    c = c & "    Dim r As Long, g As Long, b As Long" & vbCrLf
    c = c & "    r = oleColor Mod 256" & vbCrLf
    c = c & "    g = (oleColor \ 256) Mod 256" & vbCrLf
    c = c & "    b = (oleColor \ 65536) Mod 256" & vbCrLf
    c = c & "    OLEToHex = ""#"" & LCase$(Right$(""0"" & Hex$(r), 2) & Right$(""0"" & Hex$(g), 2) & Right$(""0"" & Hex$(b), 2))" & vbCrLf
    c = c & "End Function" & vbCrLf & vbCrLf
    c = c & "Private Function RegGet(ByVal section As String, ByVal key As String, ByVal def As String) As String" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    RegGet = GetSetting(APP_NAME, section, key, def)" & vbCrLf
    c = c & "End Function" & vbCrLf & vbCrLf
    c = c & "Private Sub RegSet(ByVal section As String, ByVal key As String, ByVal val As String)" & vbCrLf
    c = c & "    SaveSetting APP_NAME, section, key, val" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub LoadSettings()" & vbCrLf
    c = c & "    RowLineEnabled = CBool(RegGet(SEC_GENERAL, ""RowLineEnabled"", CStr(DEF_ROW_LINE_ENABLED)))" & vbCrLf
    c = c & "    ColLineEnabled = CBool(RegGet(SEC_GENERAL, ""ColLineEnabled"", CStr(DEF_COL_LINE_ENABLED)))" & vbCrLf
    c = c & "    RowFillEnabled = CBool(RegGet(SEC_GENERAL, ""RowFillEnabled"", CStr(DEF_ROW_FILL_ENABLED)))" & vbCrLf
    c = c & "    ColFillEnabled = CBool(RegGet(SEC_GENERAL, ""ColFillEnabled"", CStr(DEF_COL_FILL_ENABLED)))" & vbCrLf
    c = c & "    RowLineColor = HexToOLE(RegGet(SEC_GENERAL, ""RowLineColor"", DEF_ROW_LINE_COLOR))" & vbCrLf
    c = c & "    ColLineColor = HexToOLE(RegGet(SEC_GENERAL, ""ColLineColor"", DEF_COL_LINE_COLOR))" & vbCrLf
    c = c & "    RowFillColor = HexToOLE(RegGet(SEC_GENERAL, ""RowFillColor"", DEF_ROW_FILL_COLOR))" & vbCrLf
    c = c & "    ColFillColor = HexToOLE(RegGet(SEC_GENERAL, ""ColFillColor"", DEF_COL_FILL_COLOR))" & vbCrLf
    c = c & "    RowLineSize = CDbl(RegGet(SEC_GENERAL, ""RowLineSize"", CStr(DEF_ROW_LINE_SIZE)))" & vbCrLf
    c = c & "    ColLineSize = CDbl(RegGet(SEC_GENERAL, ""ColLineSize"", CStr(DEF_COL_LINE_SIZE)))" & vbCrLf
    c = c & "    RowFillOpacity = CDbl(RegGet(SEC_GENERAL, ""RowFillOpacity"", CStr(DEF_ROW_FILL_OPACITY)))" & vbCrLf
    c = c & "    ColFillOpacity = CDbl(RegGet(SEC_GENERAL, ""ColFillOpacity"", CStr(DEF_COL_FILL_OPACITY)))" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub SaveSettings()" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""RowLineEnabled"", CStr(RowLineEnabled)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""ColLineEnabled"", CStr(ColLineEnabled)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""RowFillEnabled"", CStr(RowFillEnabled)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""ColFillEnabled"", CStr(ColFillEnabled)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""RowLineColor"", OLEToHex(RowLineColor)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""ColLineColor"", OLEToHex(ColLineColor)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""RowFillColor"", OLEToHex(RowFillColor)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""ColFillColor"", OLEToHex(ColFillColor)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""RowLineSize"", CStr(RowLineSize)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""ColLineSize"", CStr(ColLineSize)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""RowFillOpacity"", CStr(RowFillOpacity)" & vbCrLf
    c = c & "    RegSet SEC_GENERAL, ""ColFillOpacity"", CStr(ColFillOpacity)" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub SaveCustomDefaults()" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""RowLineEnabled"", CStr(RowLineEnabled)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""ColLineEnabled"", CStr(ColLineEnabled)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""RowFillEnabled"", CStr(RowFillEnabled)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""ColFillEnabled"", CStr(ColFillEnabled)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""RowLineColor"", OLEToHex(RowLineColor)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""ColLineColor"", OLEToHex(ColLineColor)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""RowFillColor"", OLEToHex(RowFillColor)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""ColFillColor"", OLEToHex(ColFillColor)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""RowLineSize"", CStr(RowLineSize)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""ColLineSize"", CStr(ColLineSize)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""RowFillOpacity"", CStr(RowFillOpacity)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""ColFillOpacity"", CStr(ColFillOpacity)" & vbCrLf
    c = c & "    RegSet SEC_DEFAULTS, ""HasCustom"", ""True""" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub ResetSettings()" & vbCrLf
    c = c & "    Dim hasCustom As Boolean" & vbCrLf
    c = c & "    hasCustom = CBool(RegGet(SEC_DEFAULTS, ""HasCustom"", ""False""))" & vbCrLf
    c = c & "    If hasCustom Then" & vbCrLf
    c = c & "        RowLineEnabled = CBool(RegGet(SEC_DEFAULTS, ""RowLineEnabled"", CStr(DEF_ROW_LINE_ENABLED)))" & vbCrLf
    c = c & "        ColLineEnabled = CBool(RegGet(SEC_DEFAULTS, ""ColLineEnabled"", CStr(DEF_COL_LINE_ENABLED)))" & vbCrLf
    c = c & "        RowFillEnabled = CBool(RegGet(SEC_DEFAULTS, ""RowFillEnabled"", CStr(DEF_ROW_FILL_ENABLED)))" & vbCrLf
    c = c & "        ColFillEnabled = CBool(RegGet(SEC_DEFAULTS, ""ColFillEnabled"", CStr(DEF_COL_FILL_ENABLED)))" & vbCrLf
    c = c & "        RowLineColor = HexToOLE(RegGet(SEC_DEFAULTS, ""RowLineColor"", DEF_ROW_LINE_COLOR))" & vbCrLf
    c = c & "        ColLineColor = HexToOLE(RegGet(SEC_DEFAULTS, ""ColLineColor"", DEF_COL_LINE_COLOR))" & vbCrLf
    c = c & "        RowFillColor = HexToOLE(RegGet(SEC_DEFAULTS, ""RowFillColor"", DEF_ROW_FILL_COLOR))" & vbCrLf
    c = c & "        ColFillColor = HexToOLE(RegGet(SEC_DEFAULTS, ""ColFillColor"", DEF_COL_FILL_COLOR))" & vbCrLf
    c = c & "        RowLineSize = CDbl(RegGet(SEC_DEFAULTS, ""RowLineSize"", CStr(DEF_ROW_LINE_SIZE)))" & vbCrLf
    c = c & "        ColLineSize = CDbl(RegGet(SEC_DEFAULTS, ""ColLineSize"", CStr(DEF_COL_LINE_SIZE)))" & vbCrLf
    c = c & "        RowFillOpacity = CDbl(RegGet(SEC_DEFAULTS, ""RowFillOpacity"", CStr(DEF_ROW_FILL_OPACITY)))" & vbCrLf
    c = c & "        ColFillOpacity = CDbl(RegGet(SEC_DEFAULTS, ""ColFillOpacity"", CStr(DEF_COL_FILL_OPACITY)))" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        RowLineEnabled = DEF_ROW_LINE_ENABLED" & vbCrLf
    c = c & "        ColLineEnabled = DEF_COL_LINE_ENABLED" & vbCrLf
    c = c & "        RowFillEnabled = DEF_ROW_FILL_ENABLED" & vbCrLf
    c = c & "        ColFillEnabled = DEF_COL_FILL_ENABLED" & vbCrLf
    c = c & "        RowLineColor = HexToOLE(DEF_ROW_LINE_COLOR)" & vbCrLf
    c = c & "        ColLineColor = HexToOLE(DEF_COL_LINE_COLOR)" & vbCrLf
    c = c & "        RowFillColor = HexToOLE(DEF_ROW_FILL_COLOR)" & vbCrLf
    c = c & "        ColFillColor = HexToOLE(DEF_COL_FILL_COLOR)" & vbCrLf
    c = c & "        RowLineSize = DEF_ROW_LINE_SIZE" & vbCrLf
    c = c & "        ColLineSize = DEF_COL_LINE_SIZE" & vbCrLf
    c = c & "        RowFillOpacity = DEF_ROW_FILL_OPACITY" & vbCrLf
    c = c & "        ColFillOpacity = DEF_COL_FILL_OPACITY" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    SaveSettings" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub ShowSettings()" & vbCrLf
    c = c & "    frmSettings.Show vbModal" & vbCrLf
    c = c & "    InitEventHandler" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub InitEventHandler()" & vbCrLf
    c = c & "    If gAppEvents Is Nothing Then Set gAppEvents = New clsAppEvents" & vbCrLf
    c = c & "    Set gAppEvents.App = Application" & vbCrLf
    c = c & "End Sub" & vbCrLf
    GetModSettingsCode = c
End Function

Private Function GetModHighlighterCode() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf & vbCrLf
    c = c & "Public Const SHAPE_PREFIX As String = ""RH_""" & vbCrLf & vbCrLf
    c = c & "Private mLastRow As Long" & vbCrLf
    c = c & "Private mLastCol As Long" & vbCrLf
    c = c & "Private mLastRowCount As Long" & vbCrLf
    c = c & "Private mLastColCount As Long" & vbCrLf
    c = c & "Private mLastSheet As String" & vbCrLf & vbCrLf
    c = c & "Private mSheet As Worksheet" & vbCrLf
    c = c & "Private mRowFill As Shape" & vbCrLf
    c = c & "Private mColFill As Shape" & vbCrLf
    c = c & "Private mRowLineTop As Shape" & vbCrLf
    c = c & "Private mRowLineBot As Shape" & vbCrLf
    c = c & "Private mColLineLeft As Shape" & vbCrLf
    c = c & "Private mColLineRight As Shape" & vbCrLf & vbCrLf
    c = c & "Private Function GetOrCreateRect(ByVal ws As Worksheet, ByVal sName As String) As Shape" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set GetOrCreateRect = ws.Shapes(sName)" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If GetOrCreateRect Is Nothing Then" & vbCrLf
    c = c & "        Set GetOrCreateRect = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)" & vbCrLf
    c = c & "        GetOrCreateRect.Name = sName" & vbCrLf
    c = c & "        GetOrCreateRect.Placement = xlFreeFloating" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Function" & vbCrLf & vbCrLf
    c = c & "Private Function GetOrCreateLine(ByVal ws As Worksheet, ByVal sName As String) As Shape" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Set GetOrCreateLine = ws.Shapes(sName)" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If GetOrCreateLine Is Nothing Then" & vbCrLf
    c = c & "        Set GetOrCreateLine = ws.Shapes.AddLine(0, 0, 1, 1)" & vbCrLf
    c = c & "        GetOrCreateLine.Name = sName" & vbCrLf
    c = c & "        GetOrCreateLine.Placement = xlFreeFloating" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Function" & vbCrLf & vbCrLf
    c = c & "Private Sub EnsureShapes(ByVal ws As Worksheet)" & vbCrLf
    c = c & "    If Not mSheet Is ws Then" & vbCrLf
    c = c & "        Set mSheet = ws" & vbCrLf
    c = c & "        Set mRowFill = Nothing: Set mColFill = Nothing" & vbCrLf
    c = c & "        Set mRowLineTop = Nothing: Set mRowLineBot = Nothing" & vbCrLf
    c = c & "        Set mColLineLeft = Nothing: Set mColLineRight = Nothing" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If mRowFill Is Nothing Then Set mRowFill = GetOrCreateRect(ws, SHAPE_PREFIX & ""RowFill"")" & vbCrLf
    c = c & "    If mColFill Is Nothing Then Set mColFill = GetOrCreateRect(ws, SHAPE_PREFIX & ""ColFill"")" & vbCrLf
    c = c & "    If mRowLineTop Is Nothing Then Set mRowLineTop = GetOrCreateLine(ws, SHAPE_PREFIX & ""RowLineTop"")" & vbCrLf
    c = c & "    If mRowLineBot Is Nothing Then Set mRowLineBot = GetOrCreateLine(ws, SHAPE_PREFIX & ""RowLineBot"")" & vbCrLf
    c = c & "    If mColLineLeft Is Nothing Then Set mColLineLeft = GetOrCreateLine(ws, SHAPE_PREFIX & ""ColLineLeft"")" & vbCrLf
    c = c & "    If mColLineRight Is Nothing Then Set mColLineRight = GetOrCreateLine(ws, SHAPE_PREFIX & ""ColLineRight"")" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub ClearHighlights(ByVal ws As Worksheet)" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    Application.ScreenUpdating = False" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    For i = ws.Shapes.Count To 1 Step -1" & vbCrLf
    c = c & "        If Left$(ws.Shapes(i).Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then ws.Shapes(i).Delete" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    Set mSheet = Nothing" & vbCrLf
    c = c & "    Set mRowFill = Nothing: Set mColFill = Nothing" & vbCrLf
    c = c & "    Set mRowLineTop = Nothing: Set mRowLineBot = Nothing" & vbCrLf
    c = c & "    Set mColLineLeft = Nothing: Set mColLineRight = Nothing" & vbCrLf
    c = c & "    Application.ScreenUpdating = True" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Function HasSelectionChanged(ByVal ws As Worksheet, ByVal target As Range) As Boolean" & vbCrLf
    c = c & "    Dim sn As String" & vbCrLf
    c = c & "    sn = ws.Parent.Name & ""!"" & ws.Name" & vbCrLf
    c = c & "    If target.Row = mLastRow And target.Column = mLastCol And _" & vbCrLf
    c = c & "       target.Rows.Count = mLastRowCount And target.Columns.Count = mLastColCount And _" & vbCrLf
    c = c & "       sn = mLastSheet Then" & vbCrLf
    c = c & "        HasSelectionChanged = False" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        mLastRow = target.Row: mLastCol = target.Column" & vbCrLf
    c = c & "        mLastRowCount = target.Rows.Count: mLastColCount = target.Columns.Count" & vbCrLf
    c = c & "        mLastSheet = sn" & vbCrLf
    c = c & "        HasSelectionChanged = True" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Function" & vbCrLf & vbCrLf
    c = c & "Public Sub DrawHighlights(ByVal ws As Worksheet, ByVal target As Range)" & vbCrLf
    c = c & "    On Error GoTo ErrHandler" & vbCrLf
    c = c & "    If ws.ProtectDrawingObjects Then Exit Sub" & vbCrLf
    c = c & "    Dim visRange As Range" & vbCrLf
    c = c & "    Dim rowTop As Double, rowBottom As Double, rowHeight As Double" & vbCrLf
    c = c & "    Dim colLeft As Double, colRight As Double, colWidth As Double" & vbCrLf
    c = c & "    Dim visLeft As Double, visTop As Double, visRight As Double, visBottom As Double" & vbCrLf
    c = c & "    If Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled Or _" & vbCrLf
    c = c & "            modSettings.RowFillEnabled Or modSettings.ColFillEnabled) Then" & vbCrLf
    c = c & "        HideAllShapes ws: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    With target" & vbCrLf
    c = c & "        rowTop = .Top: rowHeight = .Height: rowBottom = rowTop + rowHeight" & vbCrLf
    c = c & "        colLeft = .Left: colWidth = .Width: colRight = colLeft + colWidth" & vbCrLf
    c = c & "    End With" & vbCrLf
    c = c & "    Set visRange = Application.ActiveWindow.VisibleRange" & vbCrLf
    c = c & "    With visRange" & vbCrLf
    c = c & "        visLeft = .Left: visTop = .Top" & vbCrLf
    c = c & "        visRight = .Left + .Width: visBottom = .Top + .Height" & vbCrLf
    c = c & "    End With" & vbCrLf
    c = c & "    Application.ScreenUpdating = False" & vbCrLf
    c = c & "    EnsureShapes ws" & vbCrLf
    c = c & "    If modSettings.RowFillEnabled And modSettings.RowFillOpacity > 0 Then" & vbCrLf
    c = c & "        With mRowFill" & vbCrLf
    c = c & "            .Left = visLeft: .Top = rowTop: .Width = visRight - visLeft: .Height = rowHeight" & vbCrLf
    c = c & "            .Fill.ForeColor.RGB = modSettings.RowFillColor" & vbCrLf
    c = c & "            .Fill.Transparency = 1# - modSettings.RowFillOpacity" & vbCrLf
    c = c & "            .Line.Visible = msoFalse: .Visible = msoTrue" & vbCrLf
    c = c & "        End With" & vbCrLf
    c = c & "    Else: mRowFill.Visible = msoFalse" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If modSettings.ColFillEnabled And modSettings.ColFillOpacity > 0 Then" & vbCrLf
    c = c & "        With mColFill" & vbCrLf
    c = c & "            .Left = colLeft: .Top = visTop: .Width = colWidth: .Height = visBottom - visTop" & vbCrLf
    c = c & "            .Fill.ForeColor.RGB = modSettings.ColFillColor" & vbCrLf
    c = c & "            .Fill.Transparency = 1# - modSettings.ColFillOpacity" & vbCrLf
    c = c & "            .Line.Visible = msoFalse: .Visible = msoTrue" & vbCrLf
    c = c & "        End With" & vbCrLf
    c = c & "    Else: mColFill.Visible = msoFalse" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If modSettings.RowLineEnabled Then" & vbCrLf
    c = c & "        PosLine mRowLineTop, visLeft, rowTop, visRight, rowTop, modSettings.RowLineColor, modSettings.RowLineSize" & vbCrLf
    c = c & "        PosLine mRowLineBot, visLeft, rowBottom, visRight, rowBottom, modSettings.RowLineColor, modSettings.RowLineSize" & vbCrLf
    c = c & "    Else: mRowLineTop.Visible = msoFalse: mRowLineBot.Visible = msoFalse" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If modSettings.ColLineEnabled Then" & vbCrLf
    c = c & "        PosLine mColLineLeft, colLeft, visTop, colLeft, visBottom, modSettings.ColLineColor, modSettings.ColLineSize" & vbCrLf
    c = c & "        PosLine mColLineRight, colRight, visTop, colRight, visBottom, modSettings.ColLineColor, modSettings.ColLineSize" & vbCrLf
    c = c & "    Else: mColLineLeft.Visible = msoFalse: mColLineRight.Visible = msoFalse" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Application.ScreenUpdating = True: Exit Sub" & vbCrLf
    c = c & "ErrHandler:" & vbCrLf
    c = c & "    Application.ScreenUpdating = True" & vbCrLf
    c = c & "    Debug.Print ""RH Error: "" & Err.Description" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub PosLine(ByVal shp As Shape, ByVal x1 As Double, ByVal y1 As Double, _" & vbCrLf
    c = c & "    ByVal x2 As Double, ByVal y2 As Double, ByVal lc As Long, ByVal lw As Double)" & vbCrLf
    c = c & "    With shp" & vbCrLf
    c = c & "        .Left = IIf(x1 < x2, x1, x2): .Top = IIf(y1 < y2, y1, y2)" & vbCrLf
    c = c & "        .Width = Abs(x2 - x1): .Height = Abs(y2 - y1)" & vbCrLf
    c = c & "        .Line.ForeColor.RGB = lc: .Line.Weight = lw" & vbCrLf
    c = c & "        .Line.Visible = msoTrue: .Visible = msoTrue" & vbCrLf
    c = c & "    End With" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub HideAllShapes(ByVal ws As Worksheet)" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    EnsureShapes ws" & vbCrLf
    c = c & "    mRowFill.Visible = msoFalse: mColFill.Visible = msoFalse" & vbCrLf
    c = c & "    mRowLineTop.Visible = msoFalse: mRowLineBot.Visible = msoFalse" & vbCrLf
    c = c & "    mColLineLeft.Visible = msoFalse: mColLineRight.Visible = msoFalse" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub ToggleRowLine()" & vbCrLf
    c = c & "    Dim s As Boolean" & vbCrLf
    c = c & "    s = Not (modSettings.RowLineEnabled And modSettings.RowFillEnabled)" & vbCrLf
    c = c & "    modSettings.RowLineEnabled = s: modSettings.RowFillEnabled = s" & vbCrLf
    c = c & "    modSettings.SaveSettings: RefreshHighlight" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub ToggleColLine()" & vbCrLf
    c = c & "    Dim s As Boolean" & vbCrLf
    c = c & "    s = Not (modSettings.ColLineEnabled And modSettings.ColFillEnabled)" & vbCrLf
    c = c & "    modSettings.ColLineEnabled = s: modSettings.ColFillEnabled = s" & vbCrLf
    c = c & "    modSettings.SaveSettings: RefreshHighlight" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Public Sub ToggleAll()" & vbCrLf
    c = c & "    Dim s As Boolean" & vbCrLf
    c = c & "    s = Not (modSettings.RowLineEnabled And modSettings.ColLineEnabled)" & vbCrLf
    c = c & "    modSettings.RowLineEnabled = s: modSettings.ColLineEnabled = s" & vbCrLf
    c = c & "    modSettings.RowFillEnabled = s: modSettings.ColFillEnabled = s" & vbCrLf
    c = c & "    modSettings.SaveSettings: RefreshHighlight" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub RefreshHighlight()" & vbCrLf
    c = c & "    If ActiveSheet Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    ClearHighlights ActiveSheet" & vbCrLf
    c = c & "    If Not ActiveCell Is Nothing Then DrawHighlights ActiveSheet, ActiveCell" & vbCrLf
    c = c & "End Sub" & vbCrLf
    GetModHighlighterCode = c
End Function

Private Function GetClsAppEventsCode() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf & vbCrLf
    c = c & "Public WithEvents App As Application" & vbCrLf & vbCrLf
    c = c & "Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf
    c = c & "    On Error GoTo ErrHandler" & vbCrLf
    c = c & "    If Not TypeOf Sh Is Worksheet Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = Sh" & vbCrLf
    c = c & "    If Not modHighlighter.HasSelectionChanged(ws, Target) Then Exit Sub" & vbCrLf
    c = c & "    modHighlighter.ClearHighlights ws" & vbCrLf
    c = c & "    modHighlighter.DrawHighlights ws, Target" & vbCrLf
    c = c & "    Exit Sub" & vbCrLf
    c = c & "ErrHandler:" & vbCrLf
    c = c & "    Debug.Print ""RH Error: "" & Err.Description" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub App_SheetActivate(ByVal Sh As Object)" & vbCrLf
    c = c & "    On Error GoTo ErrHandler" & vbCrLf
    c = c & "    If Not TypeOf Sh Is Worksheet Then Exit Sub" & vbCrLf
    c = c & "    Dim ws As Worksheet: Set ws = Sh" & vbCrLf
    c = c & "    modHighlighter.ClearHighlights ws" & vbCrLf
    c = c & "    If Not ActiveCell Is Nothing Then modHighlighter.DrawHighlights ws, ActiveCell" & vbCrLf
    c = c & "    Exit Sub" & vbCrLf
    c = c & "ErrHandler:" & vbCrLf
    c = c & "    Debug.Print ""RH Error: "" & Err.Description" & vbCrLf
    c = c & "End Sub" & vbCrLf
    GetClsAppEventsCode = c
End Function

Private Function GetThisWorkbookCode() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf & vbCrLf
    c = c & "Private Sub Workbook_Open()" & vbCrLf
    c = c & "    modSettings.LoadSettings" & vbCrLf
    c = c & "    modSettings.InitEventHandler" & vbCrLf
    c = c & "    Application.OnKey ""^+r"", ""modHighlighter.ToggleRowLine""" & vbCrLf
    c = c & "    Application.OnKey ""^+c"", ""modHighlighter.ToggleColLine""" & vbCrLf
    c = c & "    Application.OnKey ""^+a"", ""modHighlighter.ToggleAll""" & vbCrLf
    c = c & "    Application.OnKey ""^+h"", ""modSettings.ShowSettings""" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf
    c = c & "    Application.OnKey ""^+r""" & vbCrLf
    c = c & "    Application.OnKey ""^+c""" & vbCrLf
    c = c & "    Application.OnKey ""^+a""" & vbCrLf
    c = c & "    Application.OnKey ""^+h""" & vbCrLf
    c = c & "    Dim wb As Workbook, ws As Worksheet" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    For Each wb In Application.Workbooks" & vbCrLf
    c = c & "        For Each ws In wb.Worksheets" & vbCrLf
    c = c & "            modHighlighter.ClearHighlights ws" & vbCrLf
    c = c & "        Next ws" & vbCrLf
    c = c & "    Next wb" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    GetThisWorkbookCode = c
End Function

Private Function GetFrmSettingsCode() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    chkRowLine.Value = modSettings.RowLineEnabled" & vbCrLf
    c = c & "    chkColLine.Value = modSettings.ColLineEnabled" & vbCrLf
    c = c & "    chkRowFill.Value = modSettings.RowFillEnabled" & vbCrLf
    c = c & "    chkColFill.Value = modSettings.ColFillEnabled" & vbCrLf
    c = c & "    txtRowLineSize.Text = CStr(modSettings.RowLineSize)" & vbCrLf
    c = c & "    txtColLineSize.Text = CStr(modSettings.ColLineSize)" & vbCrLf
    c = c & "    txtRowFillOpacity.Text = CStr(modSettings.RowFillOpacity)" & vbCrLf
    c = c & "    txtColFillOpacity.Text = CStr(modSettings.ColFillOpacity)" & vbCrLf
    c = c & "    SetBtnColor btnRowLineColor, modSettings.RowLineColor" & vbCrLf
    c = c & "    SetBtnColor btnColLineColor, modSettings.ColLineColor" & vbCrLf
    c = c & "    SetBtnColor btnRowFillColor, modSettings.RowFillColor" & vbCrLf
    c = c & "    SetBtnColor btnColFillColor, modSettings.ColFillColor" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub SetBtnColor(ByVal btn As MSForms.CommandButton, ByVal clr As Long)" & vbCrLf
    c = c & "    btn.Caption = modSettings.OLEToHex(clr)" & vbCrLf
    c = c & "    On Error Resume Next: btn.BackColor = clr: On Error GoTo 0" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub btnRowLineColor_Click(): PickColor btnRowLineColor: End Sub" & vbCrLf
    c = c & "Private Sub btnColLineColor_Click(): PickColor btnColLineColor: End Sub" & vbCrLf
    c = c & "Private Sub btnRowFillColor_Click(): PickColor btnRowFillColor: End Sub" & vbCrLf
    c = c & "Private Sub btnColFillColor_Click(): PickColor btnColFillColor: End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub PickColor(ByVal btn As MSForms.CommandButton)" & vbCrLf
    c = c & "    If ActiveWorkbook Is Nothing Then" & vbCrLf
    c = c & "        MsgBox ""Please open a workbook first."", vbExclamation: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim curClr As Long" & vbCrLf
    c = c & "    On Error Resume Next: curClr = btn.BackColor: On Error GoTo 0" & vbCrLf
    c = c & "    ActiveWorkbook.Colors(1) = curClr" & vbCrLf
    c = c & "    If Application.Dialogs(xlDialogEditColor).Show(1) Then" & vbCrLf
    c = c & "        SetBtnColor btn, ActiveWorkbook.Colors(1)" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub btnApply_Click()" & vbCrLf
    c = c & "    modSettings.RowLineEnabled = chkRowLine.Value" & vbCrLf
    c = c & "    modSettings.ColLineEnabled = chkColLine.Value" & vbCrLf
    c = c & "    modSettings.RowFillEnabled = chkRowFill.Value" & vbCrLf
    c = c & "    modSettings.ColFillEnabled = chkColFill.Value" & vbCrLf
    c = c & "    modSettings.RowLineSize = CDbl(txtRowLineSize.Text)" & vbCrLf
    c = c & "    modSettings.ColLineSize = CDbl(txtColLineSize.Text)" & vbCrLf
    c = c & "    modSettings.RowFillOpacity = CDbl(txtRowFillOpacity.Text)" & vbCrLf
    c = c & "    modSettings.ColFillOpacity = CDbl(txtColFillOpacity.Text)" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    modSettings.RowLineColor = btnRowLineColor.BackColor" & vbCrLf
    c = c & "    modSettings.ColLineColor = btnColLineColor.BackColor" & vbCrLf
    c = c & "    modSettings.RowFillColor = btnRowFillColor.BackColor" & vbCrLf
    c = c & "    modSettings.ColFillColor = btnColFillColor.BackColor" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    modSettings.SaveSettings" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    If Not ActiveCell Is Nothing Then" & vbCrLf
    c = c & "        modHighlighter.ClearHighlights ActiveSheet" & vbCrLf
    c = c & "        modHighlighter.DrawHighlights ActiveSheet, ActiveCell" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub btnSaveDefault_Click()" & vbCrLf
    c = c & "    btnApply_Click" & vbCrLf
    c = c & "    modSettings.LoadSettings" & vbCrLf
    c = c & "    modSettings.SaveCustomDefaults" & vbCrLf
    c = c & "    MsgBox ""Defaults saved!"", vbInformation" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub btnReset_Click()" & vbCrLf
    c = c & "    modSettings.ResetSettings" & vbCrLf
    c = c & "    UserForm_Initialize" & vbCrLf
    c = c & "End Sub" & vbCrLf
    GetFrmSettingsCode = c
End Function
