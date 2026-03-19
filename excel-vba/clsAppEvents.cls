Option Explicit

Public WithEvents App As Application

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrHandler

    ' Only handle Worksheets (not Chart sheets, etc.)
    If Not TypeOf Sh Is Worksheet Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh

    ' Skip if selection didn't actually change
    If Not modHighlighter.HasSelectionChanged(ws, Target) Then Exit Sub

    ' Clear old highlights, draw new ones
    modHighlighter.ClearHighlights ws
    modHighlighter.DrawHighlights ws, Target

    Exit Sub
ErrHandler:
    Debug.Print "RH SheetSelectionChange Error: " & Err.Description
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    On Error GoTo ErrHandler

    ' Redraw when switching sheets
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
