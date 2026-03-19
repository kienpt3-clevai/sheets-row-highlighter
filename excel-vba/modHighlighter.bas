Attribute VB_Name = "modHighlighter"
Option Explicit

Public Const SHAPE_PREFIX As String = "RH_"

' Track previous highlight to avoid redundant redraws
Private mLastRow As Long
Private mLastCol As Long
Private mLastRowCount As Long
Private mLastColCount As Long
Private mLastSheet As String

' Cached shape references (reuse instead of delete/recreate)
Private mSheet As Worksheet
Private mRowFill As Shape
Private mColFill As Shape
Private mRowLineTop As Shape
Private mRowLineBot As Shape
Private mColLineLeft As Shape
Private mColLineRight As Shape

' --- Get or create a named rectangle shape ---
Private Function GetOrCreateRect(ByVal ws As Worksheet, ByVal sName As String) As Shape
    On Error Resume Next
    Set GetOrCreateRect = ws.Shapes(sName)
    On Error GoTo 0
    If GetOrCreateRect Is Nothing Then
        Set GetOrCreateRect = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
        GetOrCreateRect.Name = sName
        GetOrCreateRect.Placement = xlFreeFloating
    End If
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

' --- Ensure cached shapes exist on the given sheet ---
Private Sub EnsureShapes(ByVal ws As Worksheet)
    ' If sheet changed, invalidate cache
    If Not mSheet Is ws Then
        Set mSheet = ws
        Set mRowFill = Nothing
        Set mColFill = Nothing
        Set mRowLineTop = Nothing
        Set mRowLineBot = Nothing
        Set mColLineLeft = Nothing
        Set mColLineRight = Nothing
    End If

    If mRowFill Is Nothing Then Set mRowFill = GetOrCreateRect(ws, SHAPE_PREFIX & "RowFill")
    If mColFill Is Nothing Then Set mColFill = GetOrCreateRect(ws, SHAPE_PREFIX & "ColFill")
    If mRowLineTop Is Nothing Then Set mRowLineTop = GetOrCreateLine(ws, SHAPE_PREFIX & "RowLineTop")
    If mRowLineBot Is Nothing Then Set mRowLineBot = GetOrCreateLine(ws, SHAPE_PREFIX & "RowLineBot")
    If mColLineLeft Is Nothing Then Set mColLineLeft = GetOrCreateLine(ws, SHAPE_PREFIX & "ColLineLeft")
    If mColLineRight Is Nothing Then Set mColLineRight = GetOrCreateLine(ws, SHAPE_PREFIX & "ColLineRight")
End Sub

' --- Delete all RH_* shapes from the given sheet ---
Public Sub ClearHighlights(ByVal ws As Worksheet)
    Dim i As Long
    Application.ScreenUpdating = False
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
    On Error GoTo 0
    ' Invalidate cache
    Set mSheet = Nothing
    Set mRowFill = Nothing
    Set mColFill = Nothing
    Set mRowLineTop = Nothing
    Set mRowLineBot = Nothing
    Set mColLineLeft = Nothing
    Set mColLineRight = Nothing
    Application.ScreenUpdating = True
End Sub

' --- Check if selection actually changed (avoid redundant redraws) ---
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

' --- Main entry: draw highlights for the active cell ---
Public Sub DrawHighlights(ByVal ws As Worksheet, ByVal target As Range)
    On Error GoTo ErrHandler

    ' Skip protected sheets (can't add shapes)
    If ws.ProtectDrawingObjects Then Exit Sub

    Dim visRange As Range
    Dim rowTop As Double, rowBottom As Double, rowHeight As Double
    Dim colLeft As Double, colRight As Double, colWidth As Double
    Dim visLeft As Double, visTop As Double, visRight As Double, visBottom As Double

    ' Skip if nothing enabled
    If Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled Or _
            modSettings.RowFillEnabled Or modSettings.ColFillEnabled) Then
        HideAllShapes ws
        Exit Sub
    End If

    ' Get selection geometry (covers entire selected range)
    With target
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

    ' Ensure all cached shapes exist
    EnsureShapes ws

    ' --- Row Fill ---
    If modSettings.RowFillEnabled And modSettings.RowFillOpacity > 0 Then
        With mRowFill
            .Left = visLeft
            .Top = rowTop
            .Width = visRight - visLeft
            .Height = rowHeight
            .Fill.ForeColor.RGB = modSettings.RowFillColor
            .Fill.Transparency = 1# - modSettings.RowFillOpacity
            .Line.Visible = msoFalse
            .Visible = msoTrue
        End With
    Else
        mRowFill.Visible = msoFalse
    End If

    ' --- Col Fill ---
    If modSettings.ColFillEnabled And modSettings.ColFillOpacity > 0 Then
        With mColFill
            .Left = colLeft
            .Top = visTop
            .Width = colWidth
            .Height = visBottom - visTop
            .Fill.ForeColor.RGB = modSettings.ColFillColor
            .Fill.Transparency = 1# - modSettings.ColFillOpacity
            .Line.Visible = msoFalse
            .Visible = msoTrue
        End With
    Else
        mColFill.Visible = msoFalse
    End If

    ' --- Row Lines (top + bottom) ---
    If modSettings.RowLineEnabled Then
        PositionLine mRowLineTop, visLeft, rowTop, visRight, rowTop, _
            modSettings.RowLineColor, modSettings.RowLineSize
        PositionLine mRowLineBot, visLeft, rowBottom, visRight, rowBottom, _
            modSettings.RowLineColor, modSettings.RowLineSize
    Else
        mRowLineTop.Visible = msoFalse
        mRowLineBot.Visible = msoFalse
    End If

    ' --- Col Lines (left + right) ---
    If modSettings.ColLineEnabled Then
        PositionLine mColLineLeft, colLeft, visTop, colLeft, visBottom, _
            modSettings.ColLineColor, modSettings.ColLineSize
        PositionLine mColLineRight, colRight, visTop, colRight, visBottom, _
            modSettings.ColLineColor, modSettings.ColLineSize
    Else
        mColLineLeft.Visible = msoFalse
        mColLineRight.Visible = msoFalse
    End If

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Debug.Print "RH DrawHighlights Error: " & Err.Description
End Sub

' --- Reposition and style a line shape ---
Private Sub PositionLine(ByVal shp As Shape, _
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

' --- Hide all cached shapes ---
Private Sub HideAllShapes(ByVal ws As Worksheet)
    On Error Resume Next
    EnsureShapes ws
    mRowFill.Visible = msoFalse
    mColFill.Visible = msoFalse
    mRowLineTop.Visible = msoFalse
    mRowLineBot.Visible = msoFalse
    mColLineLeft.Visible = msoFalse
    mColLineRight.Visible = msoFalse
    On Error GoTo 0
End Sub

' --- Quick toggles (callable from shortcuts) ---
Public Sub ToggleRowLine()
    Dim newState As Boolean
    newState = Not (modSettings.RowLineEnabled And modSettings.RowFillEnabled)
    modSettings.RowLineEnabled = newState
    modSettings.RowFillEnabled = newState
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Public Sub ToggleColLine()
    Dim newState As Boolean
    newState = Not (modSettings.ColLineEnabled And modSettings.ColFillEnabled)
    modSettings.ColLineEnabled = newState
    modSettings.ColFillEnabled = newState
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
