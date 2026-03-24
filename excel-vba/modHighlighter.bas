Attribute VB_Name = "modHighlighter"
Option Explicit

' Conditional Formatting approach - works with freeze panes
' CF formulas use these prefixes for identification
Private Const CF_ROW_PREFIX As String = "=AND(ROW()>="
Private Const CF_COL_PREFIX As String = "=AND(COLUMN()>="
Private Const SHAPE_PREFIX As String = "RH_"

' Track previous highlight to avoid redundant redraws
Private mLastRow As Long
Private mLastCol As Long
Private mLastRowCount As Long
Private mLastColCount As Long
Private mLastSheet As String

' --- Blend color with white to simulate opacity ---
Private Function BlendColor(ByVal baseColor As Long, ByVal opacity As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = baseColor And &HFF
    g = (baseColor \ &H100) And &HFF
    b = (baseColor \ &H10000) And &HFF
    BlendColor = RGB(CLng(255 - (255 - r) * opacity), _
                     CLng(255 - (255 - g) * opacity), _
                     CLng(255 - (255 - b) * opacity))
End Function

' --- Map line size to CF border weight ---
Private Function SizeToWeight(ByVal sz As Double) As Long
    If sz <= 1 Then
        SizeToWeight = xlHairline
    ElseIf sz <= 2 Then
        SizeToWeight = xlThin
    ElseIf sz <= 3 Then
        SizeToWeight = xlMedium
    Else
        SizeToWeight = xlThick
    End If
End Function

' --- Check if formula is one of ours ---
Private Function IsRHFormula(ByVal formula As String) As Boolean
    IsRHFormula = (Left$(formula, Len(CF_ROW_PREFIX)) = CF_ROW_PREFIX) Or _
                  (Left$(formula, Len(CF_COL_PREFIX)) = CF_COL_PREFIX)
End Function

' --- Clear all RH highlights from sheet ---
Public Sub ClearHighlights(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long

    ' Clear CF rules (check from cell A1, works because CF applied to ws.Cells)
    For i = ws.Cells(1, 1).FormatConditions.Count To 1 Step -1
        If IsRHFormula(ws.Cells(1, 1).FormatConditions(i).Formula1) Then
            ws.Cells(1, 1).FormatConditions(i).Delete
        End If
    Next i

    ' Clean up any legacy shapes from old implementation
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i

    On Error GoTo 0
End Sub

' --- Check if selection actually changed ---
Public Function HasSelectionChanged(ByVal ws As Worksheet, ByVal target As Range) As Boolean
    Dim sheetName As String
    sheetName = ws.Parent.Name & "!" & ws.Name
    If target.Row = mLastRow And target.Column = mLastCol And _
       target.Rows.Count = mLastRowCount And target.Columns.Count = mLastColCount And _
       sheetName = mLastSheet Then
        HasSelectionChanged = False
    Else
        mLastRow = target.Row
        mLastCol = target.Column
        mLastRowCount = target.Rows.Count
        mLastColCount = target.Columns.Count
        mLastSheet = sheetName
        HasSelectionChanged = True
    End If
End Function

' --- Main: draw highlights using Conditional Formatting ---
Public Sub DrawHighlights(ByVal ws As Worksheet, ByVal target As Range)
    On Error GoTo ErrHandler

    ' Skip if nothing enabled
    If Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled Or _
            modSettings.RowFillEnabled Or modSettings.ColFillEnabled) Then
        ClearHighlights ws
        Exit Sub
    End If

    Dim targetRow As Long, targetCol As Long
    Dim targetRowEnd As Long, targetColEnd As Long
    targetRow = target.Row
    targetCol = target.Column
    targetRowEnd = targetRow + target.Rows.Count - 1
    targetColEnd = targetCol + target.Columns.Count - 1

    Application.ScreenUpdating = False

    ' Clear previous highlights
    ClearHighlights ws

    ' --- Row CF (fill + borders) ---
    If modSettings.RowLineEnabled Or modSettings.RowFillEnabled Then
        Dim rowFormula As String
        rowFormula = CF_ROW_PREFIX & targetRow & ",ROW()<=" & targetRowEnd & ")"

        Dim fcRow As FormatCondition
        Set fcRow = ws.Cells.FormatConditions.Add(xlExpression, , rowFormula)
        fcRow.StopIfTrue = False

        If modSettings.RowFillEnabled And modSettings.RowFillOpacity > 0 Then
            fcRow.Interior.Color = BlendColor(modSettings.RowFillColor, modSettings.RowFillOpacity)
        End If

        If modSettings.RowLineEnabled Then
            With fcRow.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = modSettings.RowLineColor
                .Weight = SizeToWeight(modSettings.RowLineSize)
            End With
            With fcRow.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = modSettings.RowLineColor
                .Weight = SizeToWeight(modSettings.RowLineSize)
            End With
        End If
    End If

    ' --- Col CF (fill + borders) ---
    If modSettings.ColLineEnabled Or modSettings.ColFillEnabled Then
        Dim colFormula As String
        colFormula = CF_COL_PREFIX & targetCol & ",COLUMN()<=" & targetColEnd & ")"

        Dim fcCol As FormatCondition
        Set fcCol = ws.Cells.FormatConditions.Add(xlExpression, , colFormula)
        fcCol.StopIfTrue = False

        If modSettings.ColFillEnabled And modSettings.ColFillOpacity > 0 Then
            fcCol.Interior.Color = BlendColor(modSettings.ColFillColor, modSettings.ColFillOpacity)
        End If

        If modSettings.ColLineEnabled Then
            With fcCol.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = modSettings.ColLineColor
                .Weight = SizeToWeight(modSettings.ColLineSize)
            End With
            With fcCol.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = modSettings.ColLineColor
                .Weight = SizeToWeight(modSettings.ColLineSize)
            End With
        End If
    End If

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Debug.Print "RH DrawHighlights Error: " & Err.Description
End Sub

' --- Quick toggles ---
Public Sub ToggleRowLine()
    Dim newState As Boolean
    newState = Not (modSettings.RowLineEnabled Or modSettings.RowFillEnabled)
    modSettings.RowLineEnabled = newState
    modSettings.RowFillEnabled = newState
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Public Sub ToggleColLine()
    Dim newState As Boolean
    newState = Not (modSettings.ColLineEnabled Or modSettings.ColFillEnabled)
    modSettings.ColLineEnabled = newState
    modSettings.ColFillEnabled = newState
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Public Sub ToggleAll()
    Dim newState As Boolean
    newState = Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled Or _
                    modSettings.RowFillEnabled Or modSettings.ColFillEnabled)
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
