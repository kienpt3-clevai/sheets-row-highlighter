Option Explicit

' === Shapes for lines + row fill (scrollable) + CF for row fill (freeze) ===
Private Const SHAPE_PREFIX As String = "RH_"
Private Const CF_ROW_TAG As String = "=AND(ROW()>="

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
Private mRowFill As Shape

' ======== SHAPE HELPERS ========

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

Private Sub EnsureShapes(ByVal ws As Worksheet)
    If Not mSheet Is ws Then
        Set mSheet = ws
        Set mRowLineTop = Nothing
        Set mRowLineBot = Nothing
        Set mColLineLeft = Nothing
        Set mColLineRight = Nothing
        Set mRowFill = Nothing
    End If
    If mRowLineTop Is Nothing Then Set mRowLineTop = GetOrCreateLine(ws, SHAPE_PREFIX & "RowLineTop")
    If mRowLineBot Is Nothing Then Set mRowLineBot = GetOrCreateLine(ws, SHAPE_PREFIX & "RowLineBot")
    If mColLineLeft Is Nothing Then Set mColLineLeft = GetOrCreateLine(ws, SHAPE_PREFIX & "ColLineLeft")
    If mColLineRight Is Nothing Then Set mColLineRight = GetOrCreateLine(ws, SHAPE_PREFIX & "ColLineRight")
    If mRowFill Is Nothing Then Set mRowFill = GetOrCreateRect(ws, SHAPE_PREFIX & "RowFill")
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

' ======== CF FILL HELPERS (for freeze panes) ========

Private Sub ClearCFRules(ByVal ws As Worksheet)
    Dim i As Long
    Dim fc As Object
    Dim formula As String
    On Error Resume Next
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        Set fc = ws.Cells.FormatConditions(i)
        formula = ""
        formula = fc.Formula1
        If Len(formula) > 0 Then
            If Left$(formula, Len(CF_ROW_TAG)) = CF_ROW_TAG Then
                fc.Delete
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function BlendColor(ByVal baseColor As Long, ByVal opacity As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = baseColor And &HFF
    g = (baseColor \ &H100) And &HFF
    b = (baseColor \ &H10000) And &HFF
    BlendColor = RGB(CLng(255 - (255 - r) * opacity), _
                     CLng(255 - (255 - g) * opacity), _
                     CLng(255 - (255 - b) * opacity))
End Function

' ======== PUBLIC API ========

Public Sub ClearHighlights(ByVal ws As Worksheet)
    ' Clear CF fill rules
    ClearCFRules ws

    ' Clear shapes
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
    Set mRowFill = Nothing
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
    If Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled Or _
            modSettings.RowFillEnabled Or modSettings.ColFillEnabled) Then
        ClearHighlights ws
        Exit Sub
    End If

    Dim targetRow As Long, targetRowEnd As Long
    targetRow = target.Row
    targetRowEnd = targetRow + target.Rows.Count - 1

    Application.ScreenUpdating = False

    ' === CF Row Fill (freeze panes) ===
    ' Uses RowFill settings (checkbox Row Fill, color, opacity)
    ClearCFRules ws
    If modSettings.RowFillEnabled And modSettings.RowFillOpacity > 0 Then
        On Error Resume Next
        Dim rowF As String
        rowF = CF_ROW_TAG & targetRow & ",ROW()<=" & targetRowEnd & ")"
        Dim fcR As FormatCondition
        Set fcR = ws.Cells.FormatConditions.Add(xlExpression, , rowF)
        If Not fcR Is Nothing Then
            fcR.StopIfTrue = False
            fcR.Interior.Color = BlendColor(modSettings.RowFillColor, modSettings.RowFillOpacity)
        End If
        On Error GoTo 0
    End If

    ' === Shape drawing ===
    On Error GoTo ErrHandler

    If ws.ProtectDrawingObjects Then GoTo Done

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

    Dim rowTop As Double, rowBottom As Double, rowHeight As Double
    Dim colLeft As Double, colRight As Double
    With target
        rowTop = .Top
        rowHeight = .Height
        rowBottom = rowTop + rowHeight
        colLeft = .Left
        colRight = colLeft + .Width
    End With

    EnsureShapes ws

    ' --- Shape Row Fill (scrollable) ---
    ' Uses ColFill settings (checkbox Col Fill, color, opacity) repurposed
    If modSettings.ColFillEnabled And modSettings.ColFillOpacity > 0 Then
        With mRowFill
            .Left = visLeft
            .Top = rowTop
            .Width = visRight - visLeft
            .Height = rowHeight
            .Fill.ForeColor.RGB = modSettings.ColFillColor
            .Fill.Transparency = 1# - modSettings.ColFillOpacity
            .Line.Visible = msoFalse
            .Visible = msoTrue
        End With
    Else
        mRowFill.Visible = msoFalse
    End If

    ' --- Row Lines ---
    If modSettings.RowLineEnabled Then
        PosLine mRowLineTop, visLeft, rowTop, visRight, rowTop, _
            modSettings.RowLineColor, modSettings.RowLineSize
        PosLine mRowLineBot, visLeft, rowBottom, visRight, rowBottom, _
            modSettings.RowLineColor, modSettings.RowLineSize
    Else
        mRowLineTop.Visible = msoFalse
        mRowLineBot.Visible = msoFalse
    End If

    ' --- Col Lines ---
    If modSettings.ColLineEnabled Then
        PosLine mColLineLeft, colLeft, visTop, colLeft, visBottom, _
            modSettings.ColLineColor, modSettings.ColLineSize
        PosLine mColLineRight, colRight, visTop, colRight, visBottom, _
            modSettings.ColLineColor, modSettings.ColLineSize
    Else
        mColLineLeft.Visible = msoFalse
        mColLineRight.Visible = msoFalse
    End If

Done:
    ' Replace undo stack entries from shape/CF ops with a no-op
    ' so Ctrl+Z doesn't undo highlight operations
    On Error Resume Next
    Application.OnUndo "", "modHighlighter.NoOp"
    On Error GoTo 0

    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
End Sub

' --- No-op for undo stack ---
Public Sub NoOp()
    ' Intentionally empty - absorbs one Ctrl+Z press
End Sub

Private Sub HideAllShapes(ByVal ws As Worksheet)
    On Error Resume Next
    EnsureShapes ws
    mRowFill.Visible = msoFalse
    mRowLineTop.Visible = msoFalse
    mRowLineBot.Visible = msoFalse
    mColLineLeft.Visible = msoFalse
    mColLineRight.Visible = msoFalse
    On Error GoTo 0
End Sub

' ======== TOGGLES ========

Public Sub ToggleRowLine()
    Dim s As Boolean
    s = Not (modSettings.RowLineEnabled Or modSettings.RowFillEnabled)
    modSettings.RowLineEnabled = s
    modSettings.RowFillEnabled = s
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Public Sub ToggleColLine()
    Dim s As Boolean
    s = Not (modSettings.ColLineEnabled Or modSettings.ColFillEnabled)
    modSettings.ColLineEnabled = s
    modSettings.ColFillEnabled = s
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Public Sub ToggleAll()
    Dim s As Boolean
    s = Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled Or _
             modSettings.RowFillEnabled Or modSettings.ColFillEnabled)
    modSettings.RowLineEnabled = s
    modSettings.ColLineEnabled = s
    modSettings.RowFillEnabled = s
    modSettings.ColFillEnabled = s
    modSettings.SaveSettings
    RefreshHighlight
End Sub

Public Sub RefreshHighlight()
    If ActiveSheet Is Nothing Then Exit Sub
    ClearHighlights ActiveSheet
    If Not ActiveCell Is Nothing Then
        DrawHighlights ActiveSheet, ActiveCell
    End If
End Sub
