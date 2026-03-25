Attribute VB_Name = "modHighlighter"
Option Explicit

Private Const SHAPE_PREFIX As String = "RH_"

Private mLastRow As Long
Private mLastCol As Long
Private mLastRowCount As Long
Private mLastColCount As Long
Private mLastSheet As String

Private mSheet As Worksheet
Private mRowLineTop As Shape
Private mRowLineBot As Shape
Private mColLineLeft As Shape
Private mColLineRight As Shape

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

Private Sub EnsureShapes(ByVal ws As Worksheet)
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

Public Sub ClearHighlights(ByVal ws As Worksheet)
    On Error Resume Next
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
    On Error GoTo 0
    Set mSheet = Nothing
    Set mRowLineTop = Nothing
    Set mRowLineBot = Nothing
    Set mColLineLeft = Nothing
    Set mColLineRight = Nothing
End Sub

Public Function HasSelectionChanged(ByVal ws As Worksheet, ByVal target As Range) As Boolean
    Dim sn As String
    sn = ws.Parent.Name & "!" & ws.Name
    If target.Row = mLastRow And target.Column = mLastCol And _
       target.Rows.Count = mLastRowCount And target.Columns.Count = mLastColCount And _
       sn = mLastSheet Then
        HasSelectionChanged = False
    Else
        mLastRow = target.Row
        mLastCol = target.Column
        mLastRowCount = target.Rows.Count
        mLastColCount = target.Columns.Count
        mLastSheet = sn
        HasSelectionChanged = True
    End If
End Function

Public Sub DrawHighlights(ByVal ws As Worksheet, ByVal target As Range)
    On Error GoTo ErrHandler

    If ws.ProtectDrawingObjects Then Exit Sub

    If Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled) Then
        HideAllShapes ws
        Exit Sub
    End If

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

    Dim rowTop As Double, rowBottom As Double
    Dim colLeft As Double, colRight As Double
    With target
        rowTop = .Top
        rowBottom = rowTop + .Height
        colLeft = .Left
        colRight = colLeft + .Width
    End With

    Application.ScreenUpdating = False
    EnsureShapes ws

    If modSettings.RowLineEnabled Then
        PosLine mRowLineTop, visLeft, rowTop, visRight, rowTop, _
            modSettings.RowLineColor, modSettings.RowLineSize
        PosLine mRowLineBot, visLeft, rowBottom, visRight, rowBottom, _
            modSettings.RowLineColor, modSettings.RowLineSize
    Else
        mRowLineTop.Visible = msoFalse
        mRowLineBot.Visible = msoFalse
    End If

    If modSettings.ColLineEnabled Then
        PosLine mColLineLeft, colLeft, visTop, colLeft, visBottom, _
            modSettings.ColLineColor, modSettings.ColLineSize
        PosLine mColLineRight, colRight, visTop, colRight, visBottom, _
            modSettings.ColLineColor, modSettings.ColLineSize
    Else
        mColLineLeft.Visible = msoFalse
        mColLineRight.Visible = msoFalse
    End If

    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
End Sub

Private Sub HideAllShapes(ByVal ws As Worksheet)
    On Error Resume Next
    EnsureShapes ws
    mRowLineTop.Visible = msoFalse
    mRowLineBot.Visible = msoFalse
    mColLineLeft.Visible = msoFalse
    mColLineRight.Visible = msoFalse
    On Error GoTo 0
End Sub

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
    newState = Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled)
    modSettings.RowLineEnabled = newState
    modSettings.ColLineEnabled = newState
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
