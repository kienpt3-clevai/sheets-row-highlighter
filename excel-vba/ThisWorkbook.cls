VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================
' This code goes into Personal.xlsb > ThisWorkbook
' It initializes the row highlighter on Excel startup
' ============================================================
Option Explicit

Private Sub Workbook_Open()
    modSettings.LoadSettings
    modSettings.InitEventHandler

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
