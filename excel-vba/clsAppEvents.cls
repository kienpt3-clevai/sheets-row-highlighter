Option Explicit

Public WithEvents App As Application

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrHandler
    If Not TypeOf Sh Is Worksheet Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh

    If Not modHighlighter.HasSelectionChanged(ws, Target) Then Exit Sub

    modHighlighter.ClearHighlights ws
    modHighlighter.DrawHighlights ws, Target

    Exit Sub
ErrHandler:
    Debug.Print "RH SheetSelectionChange Error: " & Err.Description
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    On Error GoTo ErrHandler
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Dim ws As Worksheet
    Set ws = Sh
    modHighlighter.ClearHighlights ws
    If Not ActiveCell Is Nothing Then
        modHighlighter.DrawHighlights ws, ActiveCell
    End If
    Exit Sub
ErrHandler:
    Debug.Print "RH SheetActivate Error: " & Err.Description
End Sub

' --- Clean highlights BEFORE save so shared files stay clean ---
Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In Wb.Worksheets
        modHighlighter.ClearHighlights ws
    Next ws
    On Error GoTo 0
End Sub

' --- Redraw AFTER save completes ---
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    On Error Resume Next
    If Success Then
        If Not ActiveCell Is Nothing Then
            If TypeOf ActiveSheet Is Worksheet Then
                modHighlighter.DrawHighlights ActiveSheet, ActiveCell
            End If
        End If
    End If
    On Error GoTo 0
End Sub

' --- Clean when switching away from a workbook ---
Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In Wb.Worksheets
        modHighlighter.ClearHighlights ws
    Next ws
    On Error GoTo 0
End Sub

' --- Redraw when coming back to a workbook ---
Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    If Not ActiveCell Is Nothing Then
        If TypeOf ActiveSheet Is Worksheet Then
            modHighlighter.DrawHighlights ActiveSheet, ActiveCell
        End If
    End If
    On Error GoTo 0
End Sub
