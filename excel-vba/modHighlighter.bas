Option Explicit

' === Shapes for lines + CF for fills (fills work in freeze panes) ===
Private Const SHAPE_PREFIX As String = "RH_"
Private Const CF_ROW_TAG As String = "=AND(ROW()>="
Private Const CF_COL_TAG As String = "=AND(COLUMN()>="

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

' ======== CF FILL HELPERS ========

Private Function BlendColor(ByVal baseColor As Long, ByVal opacity As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = baseColor And &HFF
    g = (baseColor \ &H100) And &HFF
    b = (baseColor \ &H10000) And &HFF
    BlendColor = RGB(CLng(255 - (255 - r) * opacity), _
                     CLng(255 - (255 - g) * opacity), _
                     CLng(255 - (255 - b) * opacity))
End Function

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
            If Left$(formula, Len(CF_ROW_TAG)) = CF_ROW_TAG Or _
               Left$(formula, Len(CF_COL_TAG)) = CF_COL_TAG Then
                fc.Delete
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub DrawCFFills(ByVal ws As Worksheet, ByVal target As Range)
    Dim targetRow As Long, targetRowEnd As Long
    Dim targetCol As Long, targetColEnd As Long
    targetRow = target.Row
    targetCol = target.Column
    targetRowEnd = targetRow + target.Rows.Count - 1
    targetColEnd = targetCol + target.Columns.Count - 1

    On Error Resume Next

    If modSettings.RowFillEnabled And modSettings.RowFillOpacity > 0 Then
        Dim rowF As String
        rowF = CF_ROW_TAG & targetRow & ",ROW()<=" & targetRowEnd & ")"
        Dim fcR As FormatCondition
        Set fcR = ws.Cells.FormatConditions.Add(xlExpression, , rowF)
        If Not fcR Is Nothing Then
            fcR.StopIfTrue = False
            fcR.Interior.Color = BlendColor(modSettings.RowFillColor, modSettings.RowFillOpacity)
        End If
    End If

    If modSettings.ColFillEnabled And modSettings.ColFillOpacity > 0 Then
        Dim colF As String
        colF = CF_COL_TAG & targetCol & ",COLUMN()<=" & targetColEnd & ")"
        Dim fcC As FormatCondition
        Set fcC = ws.Cells.FormatConditions.Add(xlExpression, , colF)
        If Not fcC Is Nothing Then
            fcC.StopIfTrue = False
            fcC.Interior.Color = BlendColor(modSettings.ColFillColor, modSettings.ColFillOpacity)
        End If
    End If

    On Error GoTo 0
End Sub

' ======== PUBLIC API ========

Public Sub ClearHighlights(ByVal ws As Worksheet)
    ' Clear CF fills (separate error handling)
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

    Application.ScreenUpdating = False

    ' --- CF fills (tach rieng, loi khong anh huong shapes) ---
    ClearCFRules ws
    If modSettings.RowFillEnabled Or modSettings.ColFillEnabled Then
        DrawCFFills ws, target
    End If

    ' --- Shape lines ---
    On Error GoTo LineErr

    If Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled) Then
        HideAllShapes ws
        GoTo Done
    End If

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

    Dim rowTop As Double, rowBottom As Double
    Dim colLeft As Double, colRight As Double
    With target
        rowTop = .Top
        rowBottom = rowTop + .Height
        colLeft = .Left
        colRight = colLeft + .Width
    End With

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

Done:
    Application.ScreenUpdating = True
    Exit Sub
LineErr:
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

Private Sub RefreshHighlight()
    If ActiveSheet Is Nothing Then Exit Sub
    ClearHighlights ActiveSheet
    If Not ActiveCell Is Nothing Then
        DrawHighlights ActiveSheet, ActiveCell
    End If
End Sub
