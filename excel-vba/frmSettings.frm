VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings
   Caption         =   "Excel Row Highlighter Settings"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' frmSettings - Excel Row Highlighter Settings UserForm
'
' NOTE: The OleObjectBlob above references a .frx file that
' stores the binary layout of controls. Since this cannot be
' generated from text, you must CREATE THE FORM MANUALLY in
' VBA Editor and then paste this code into it.
'
' Required controls (create in VBA Editor Form Designer):
'   CheckBox:      chkRowLine, chkColLine, chkRowFill, chkColFill
'   TextBox:       txtRowLineSize, txtColLineSize, txtRowFillOpacity, txtColFillOpacity
'   CommandButton: btnRowLineColor, btnColLineColor, btnRowFillColor, btnColFillColor
'   CommandButton: btnReset, btnApply
'   Label:         lblRowLine, lblColLine, lblRowFill, lblColFill
'                  lblRowLineSize, lblColLineSize, lblRowFillOp, lblColFillOp
'
' Layout guide:
'   Row 1: chkRowLine  | lblRowLineSize + txtRowLineSize | btnRowLineColor
'   Row 2: chkColLine  | lblColLineSize + txtColLineSize | btnColLineColor
'   Row 3: chkRowFill  | lblRowFillOp + txtRowFillOpacity | btnRowFillColor
'   Row 4: chkColFill  | lblColFillOp + txtColFillOpacity | btnColFillColor
'   Row 5: btnReset    |                                  | btnApply
' ============================================================
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

' --- Color picker buttons ---
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

' --- Apply button ---
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
    On Error Resume Next
    If Not ActiveCell Is Nothing Then
        modHighlighter.ClearHighlights ActiveSheet
        modHighlighter.DrawHighlights ActiveSheet, ActiveCell
    End If
    On Error GoTo 0

    Unload Me
End Sub

' --- Reset button ---
Private Sub btnReset_Click()
    modSettings.ResetSettings
    UserForm_Initialize  ' Refresh controls with defaults
End Sub
