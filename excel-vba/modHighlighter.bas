Attribute VB_Name = "modHighlighter"
Option Explicit

' === CF for fills (works in freeze panes), Shapes for lines ===
Private Const CF_ROW_PREFIX As String = "=AND(ROW()>="
Private Const CF_COL_PREFIX As String = "=AND(COLUMN()>="
Private Const SHAPE_PREFIX As String = "RH_"

' Track previous highlight
Private mLastRow As Long
Private mLastCol As Long
Private mLastRowCount As Long
Private mLastColCount As Long
Private mLastSheet As String

' Cached line shapes
Private mSheet As Worksheet
Private mRowLineTop As Shape
Private mRowLineBot As Shape
Private mColLineLeft As Shape
Private mColLineRight As Shape

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

' --- Check if formula is one of ours ---
Private Function IsRHFormula(ByVal formula As String) As Boolean
    IsRHFormula = (Left$(formula, Len(CF_ROW_PREFIX)) = CF_ROW_PREFIX) Or _
                  (Left$(formula, Len(CF_COL_PREFIX)) = CF_COL_PREFIX)
End Function

' --- Get or create a named line shape ---
Private Function GetOrCreateLine(ByVal ws As Worksheet, ByVal sName As String) As Shape
    On Error Resume Next
    Set GetOrCreateLine = ws.Shapes(sName)
    On Error GoTo 0
    If GetOrCreateLine Is Nothing Then
        Set GetOrCreateLine = ws.Shapes.AddLine(0, 0, 1, 1)
        GetOrCreateLine.Name = sName
        GetOrCreateLine.Placement = xlFreeFloating
    End If
End Function

' --- Ensure cached line shapes exist ---
Private Sub EnsureLineShapes(ByVal ws As Worksheet)
    If Not mSheet Is ws Then
        Set mSheet = ws
        Set mRowLineTop = Nothing
        Set mRowLineBot = Nothing
        Set mColLineLeft = Nothing
        Set mColLineRight = Nothing
    End If
    If mRowLineTop Is Nothing Then Set mRowLineTop = GetOrCreateLine(ws, SHAPE_PREFIX & "RowLineTop")
    If mRowLineBot Is Nothing Then Set mRowLineBot = GetOrCreateLine(ws, SHAPE_PREFIX & "RowLineBot")
    If mColLineLeft Is Nothing Then Set mColLineLeft = GetOrCreateLine(ws, SHAPE_PREFIX & "ColLineLeft")
    If mColLineRight Is Nothing Then Set mColLineRight = GetOrCreateLine(ws, SHAPE_PREFIX & "ColLineRight")
End Sub

' --- Position a line shape ---
Private Sub PosLine(ByVal shp As Shape, _
    ByVal x1 As Double, ByVal y1 As Double, _
    ByVal x2 As Double, ByVal y2 As Double, _
    ByVal lineColor As Long, ByVal lineWeight As Double)
    With shp
        .Left = IIf(x1 < x2, x1, x2)
        .Top = IIf(y1 < y2, y1, y2)
        .Width = Abs(x2 - x1)
        .Height = Abs(y2 - y1)
        .Line.ForeColor.RGB = lineColor
        .Line.Weight = lineWeight
        .Line.Visible = msoTrue
        .Visible = msoTrue
    End With
End Sub

' --- Clear all RH highlights from sheet ---
Public Sub ClearHighlights(ByVal ws As Worksheet)
    Dim i As Long

    ' Clear CF fill rules
    On Error Resume Next
    For i = ws.Cells(1, 1).FormatConditions.Count To 1 Step -1
        Dim f As String
        f = ws.Cells(1, 1).FormatConditions(i).Formula1
        If IsRHFormula(f) Then
            ws.Cells(1, 1).FormatConditions(i).Delete
        End If
    Next i
    On Error GoTo 0

    ' Clear/hide line shapes
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
    On Error GoTo 0

    ' Invalidate shape cache
    Set mSheet = Nothing
    Set mRowLineTop = Nothing
    Set mRowLineBot = Nothing
    Set mColLineLeft = Nothing
    Set mColLineRight = Nothing
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

' --- Main: draw highlights ---
Public Sub DrawHighlights(ByVal ws As Worksheet, ByVal target As Range)
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

    ' === STEP 1: Clear old CF rules ===
    On Error Resume Next
    Dim i As Long
    For i = ws.Cells(1, 1).FormatConditions.Count To 1 Step -1
        Dim f As String
        f = ws.Cells(1, 1).FormatConditions(i).Formula1
        If IsRHFormula(f) Then
            ws.Cells(1, 1).FormatConditions(i).Delete
        End If
    Next i
    On Error GoTo 0

    ' === STEP 2: Add CF fills ===
    On Error Resume Next

    ' Row fill
    If modSettings.RowFillEnabled And modSettings.RowFillOpacity > 0 Then
        Dim rowFormula As String
        rowFormula = CF_ROW_PREFIX & targetRow & ",ROW()<=" & targetRowEnd & ")"
        Dim fcRow As FormatCondition
        Set fcRow = ws.Cells.FormatConditions.Add(xlExpression, , rowFormula)
        If Not fcRow Is Nothing Then
            fcRow.StopIfTrue = False
            fcRow.Interior.Color = BlendColor(modSettings.RowFillColor, modSettings.RowFillOpacity)
        End If
    End If

    ' Col fill
    If modSettings.ColFillEnabled And modSettings.ColFillOpacity > 0 Then
        Dim colFormula As String
        colFormula = CF_COL_PREFIX & targetCol & ",COLUMN()<=" & targetColEnd & ")"
        Dim fcCol As FormatCondition
        Set fcCol = ws.Cells.FormatConditions.Add(xlExpression, , colFormula)
        If Not fcCol Is Nothing Then
            fcCol.StopIfTrue = False
            fcCol.Interior.Color = BlendColor(modSettings.ColFillColor, modSettings.ColFillOpacity)
        End If
    End If

    On Error GoTo 0

    ' === STEP 3: Draw line shapes ===
    On Error GoTo LineErr

    ' Get visible range for line extent
    Dim visRange As Range
    Set visRange = Application.ActiveWindow.VisibleRange
    Dim area As Range
    Dim visLeft As Double, visTop As Double, visRight As Double, visBottom As Double
    visLeft = 9999999#: visTop = 9999999#: visRight = 0#: visBottom = 0#
    For Each area In visRange.Areas
        If area.Left < visLeft Then visLeft = area.Left
        If area.Top < visTop Then visTop = area.Top
        If area.Left + area.Width > visRight Then visRight = area.Left + area.Width
        If area.Top + area.Height > visBottom Then visBottom = area.Top + area.Height
    Next area

    ' Cell geometry
    Dim rowTop As Double, rowBottom As Double, rowHeight As Double
    Dim colLeft As Double, colRight As Double, colWidth As Double
    With target
        rowTop = .Top: rowHeight = .Height: rowBottom = rowTop + rowHeight
        colLeft = .Left: colWidth = .Width: colRight = colLeft + colWidth
    End With

    If ws.ProtectDrawingObjects Then GoTo SkipLines

    EnsureLineShapes ws

    ' Row lines
    If modSettings.RowLineEnabled Then
        PosLine mRowLineTop, visLeft, rowTop, visRight, rowTop, _
            modSettings.RowLineColor, modSettings.RowLineSize
        PosLine mRowLineBot, visLeft, rowBottom, visRight, rowBottom, _
            modSettings.RowLineColor, modSettings.RowLineSize
    Else
        mRowLineTop.Visible = msoFalse
        mRowLineBot.Visible = msoFalse
    End If

    ' Col lines
    If modSettings.ColLineEnabled Then
        PosLine mColLineLeft, colLeft, visTop, colLeft, visBottom, _
            modSettings.ColLineColor, modSettings.ColLineSize
        PosLine mColLineRight, colRight, visTop, colRight, visBottom, _
            modSettings.ColLineColor, modSettings.ColLineSize
    Else
        mColLineLeft.Visible = msoFalse
        mColLineRight.Visible = msoFalse
    End If

SkipLines:
    Application.ScreenUpdating = True
    Exit Sub

LineErr:
    Application.ScreenUpdating = True
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
