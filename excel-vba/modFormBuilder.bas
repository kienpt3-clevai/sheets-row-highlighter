Attribute VB_Name = "modFormBuilder"
Option Explicit

' ============================================================
' Run this ONCE to create frmSettings UserForm automatically.
' After running, delete this module from PERSONAL.XLSB.
'
' PREREQUISITE: File > Options > Trust Center > Trust Center Settings
'   > Macro Settings > Tick "Trust access to the VBA project object model"
' ============================================================
Public Sub BuildSettingsForm()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim frm As Object
    Dim ctrl As Object
    Dim codeModule As Object

    ' Get PERSONAL.XLSB VBA project
    Set vbProj = Workbooks("PERSONAL.XLSB").VBProject

    ' Remove any existing frmSettings or leftover UserForm
    Dim existComp As Object
    For Each existComp In vbProj.VBComponents
        If existComp.Type = 3 Then ' vbext_ct_MSForm
            If existComp.Name = "frmSettings" Or Left$(existComp.Name, 8) = "UserForm" Then
                vbProj.VBComponents.Remove existComp
            End If
        End If
    Next existComp

    ' Create UserForm (3 = vbext_ct_MSForm)
    Set vbComp = vbProj.VBComponents.Add(3)

    ' Set form properties via Properties collection
    vbComp.Properties("Caption") = "Excel Row Highlighter Settings"
    vbComp.Properties("Width") = 340
    vbComp.Properties("Height") = 220

    Set frm = vbComp.Designer

    Dim yPos As Single
    Dim rowH As Single
    rowH = 30
    yPos = 10

    ' === Row 1: Row Line ===
    MakeCheckBox frm, "chkRowLine", "Row Line", 10, yPos, 75, 18
    MakeLabel frm, "lblRowLineSize", "Size:", 95, yPos + 2, 30, 14
    MakeTextBox frm, "txtRowLineSize", "3.25", 127, yPos, 45, 18
    MakeButton frm, "btnRowLineColor", "#c2185b", 185, yPos, 135, 18
    yPos = yPos + rowH

    ' === Row 2: Col Line ===
    MakeCheckBox frm, "chkColLine", "Col Line", 10, yPos, 75, 18
    MakeLabel frm, "lblColLineSize", "Size:", 95, yPos + 2, 30, 14
    MakeTextBox frm, "txtColLineSize", "3.25", 127, yPos, 45, 18
    MakeButton frm, "btnColLineColor", "#c2185b", 185, yPos, 135, 18
    yPos = yPos + rowH

    ' === Row 3: Row Fill ===
    MakeCheckBox frm, "chkRowFill", "Row Fill", 10, yPos, 75, 18
    MakeLabel frm, "lblRowFillOp", "Opacity:", 95, yPos + 2, 40, 14
    MakeTextBox frm, "txtRowFillOpacity", "0.05", 137, yPos, 35, 18
    MakeButton frm, "btnRowFillColor", "#c2185b", 185, yPos, 135, 18
    yPos = yPos + rowH

    ' === Row 4: Col Fill ===
    MakeCheckBox frm, "chkColFill", "Col Fill", 10, yPos, 75, 18
    MakeLabel frm, "lblColFillOp", "Opacity:", 95, yPos + 2, 40, 14
    MakeTextBox frm, "txtColFillOpacity", "0.05", 137, yPos, 35, 18
    MakeButton frm, "btnColFillColor", "#c2185b", 185, yPos, 135, 18
    yPos = yPos + rowH + 10

    ' === Row 5: Action Buttons ===
    MakeButton frm, "btnReset", "Reset Defaults", 10, yPos, 100, 26
    MakeButton frm, "btnApply", "Apply && Close", 220, yPos, 100, 26

    ' === Add code behind ===
    Set codeModule = vbComp.codeModule
    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If

    Dim c As String
    c = ""
    c = c & "Option Explicit" & vbCrLf & vbCrLf
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
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    btn.BackColor = clr" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub btnRowLineColor_Click()" & vbCrLf
    c = c & "    PickColor btnRowLineColor" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub btnColLineColor_Click()" & vbCrLf
    c = c & "    PickColor btnColLineColor" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub btnRowFillColor_Click()" & vbCrLf
    c = c & "    PickColor btnRowFillColor" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub btnColFillColor_Click()" & vbCrLf
    c = c & "    PickColor btnColFillColor" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub PickColor(ByVal btn As MSForms.CommandButton)" & vbCrLf
    c = c & "    If ActiveWorkbook Is Nothing Then" & vbCrLf
    c = c & "        MsgBox ""Please open a workbook first."", vbExclamation" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Dim curClr As Long" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    curClr = btn.BackColor" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
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
    c = c & "Private Sub btnReset_Click()" & vbCrLf
    c = c & "    modSettings.ResetSettings" & vbCrLf
    c = c & "    UserForm_Initialize" & vbCrLf
    c = c & "End Sub" & vbCrLf

    codeModule.AddFromString c

    ' Try to rename the form
    On Error Resume Next
    vbComp.Name = "frmSettings"
    On Error GoTo 0

    Dim finalName As String
    finalName = vbComp.Name

    If finalName = "frmSettings" Then
        MsgBox "frmSettings created successfully!" & vbCrLf & vbCrLf & _
               "Press Ctrl+Shift+H to open settings." & vbCrLf & _
               "You can delete modFormBuilder now.", vbInformation
    Else
        MsgBox "Form created as '" & finalName & "'." & vbCrLf & vbCrLf & _
               "Please rename it to 'frmSettings':" & vbCrLf & _
               "1. Click the form in Project Explorer" & vbCrLf & _
               "2. In Properties window, change (Name) to: frmSettings" & vbCrLf & _
               "3. Press Enter", vbInformation
    End If
End Sub

' === Helper: Add CheckBox ===
Private Sub MakeCheckBox(ByVal frm As Object, ByVal sName As String, ByVal sCaption As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim ctrl As Object
    Set ctrl = frm.Controls.Add("Forms.CheckBox.1", sName)
    ctrl.Left = x
    ctrl.Top = y
    ctrl.Width = w
    ctrl.Height = h
    ctrl.Caption = sCaption
    ctrl.Value = True
End Sub

' === Helper: Add Label ===
Private Sub MakeLabel(ByVal frm As Object, ByVal sName As String, ByVal sCaption As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim ctrl As Object
    Set ctrl = frm.Controls.Add("Forms.Label.1", sName)
    ctrl.Left = x
    ctrl.Top = y
    ctrl.Width = w
    ctrl.Height = h
    ctrl.Caption = sCaption
End Sub

' === Helper: Add TextBox ===
Private Sub MakeTextBox(ByVal frm As Object, ByVal sName As String, ByVal sText As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim ctrl As Object
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", sName)
    ctrl.Left = x
    ctrl.Top = y
    ctrl.Width = w
    ctrl.Height = h
    ctrl.Text = sText
End Sub

' === Helper: Add CommandButton ===
Private Sub MakeButton(ByVal frm As Object, ByVal sName As String, ByVal sCaption As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim ctrl As Object
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", sName)
    ctrl.Left = x
    ctrl.Top = y
    ctrl.Width = w
    ctrl.Height = h
    ctrl.Caption = sCaption
End Sub
