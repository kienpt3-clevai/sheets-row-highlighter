Attribute VB_Name = "modHighlighter"
Option Explicit

Public Const SHAPE_PREFIX As String = "RH_"

' Track previous highlight to avoid redundant redraws
Private mLastRow As Long
Private mLastCol As Long
Private mLastSheet As String

' --- Delete all RH_* shapes from the given sheet ---
Public Sub ClearHighlights(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim i As Long
    ' Iterate backwards to safely delete
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            shp.Delete
        End If
    Next i
End Sub

' --- Check if selection actually changed (avoid redundant redraws) ---
Public Function HasSelectionChanged(ByVal ws As Worksheet, ByVal target As Range) As Boolean
    Dim sheetName As String
    sheetName = ws.Parent.Name & "!" & ws.Name
    If target.Row = mLastRow And target.Column = mLastCol And sheetName = mLastSheet Then
        HasSelectionChanged = False
    Else
        mLastRow = target.Row
        mLastCol = target.Column
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
    Dim cellRow As Long, cellCol As Long
    Dim rowTop As Double, rowBottom As Double, rowHeight As Double
    Dim colLeft As Double, colRight As Double, colWidth As Double
    Dim visLeft As Double, visTop As Double, visRight As Double, visBottom As Double
    Dim shp As Shape

    ' Skip if nothing enabled
    If Not (modSettings.RowLineEnabled Or modSettings.ColLineEnabled Or _
            modSettings.RowFillEnabled Or modSettings.ColFillEnabled) Then
        Exit Sub
    End If

    cellRow = target.Row
    cellCol = target.Column

    ' Get active cell geometry (handle merged cells)
    Dim mergedArea As Range
    Set mergedArea = target.MergeArea
    With mergedArea
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

    ' --- Row Fill ---
    If modSettings.RowFillEnabled And modSettings.RowFillOpacity > 0 Then
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
            visLeft, rowTop, visRight - visLeft, rowHeight)
        With shp
            .Name = SHAPE_PREFIX & "RowFill"
            .Fill.ForeColor.RGB = modSettings.RowFillColor
            .Fill.Transparency = 1# - modSettings.RowFillOpacity
            .Line.Visible = msoFalse
            .Placement = xlFreeFloating
        End With
    End If

    ' --- Col Fill ---
    If modSettings.ColFillEnabled And modSettings.ColFillOpacity > 0 Then
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
            colLeft, visTop, colWidth, visBottom - visTop)
        With shp
            .Name = SHAPE_PREFIX & "ColFill"
            .Fill.ForeColor.RGB = modSettings.ColFillColor
            .Fill.Transparency = 1# - modSettings.ColFillOpacity
            .Line.Visible = msoFalse
            .Placement = xlFreeFloating
        End With
    End If

    ' --- Row Lines (top + bottom) ---
    If modSettings.RowLineEnabled Then
        ' Top line
        Set shp = ws.Shapes.AddLine( _
            visLeft, rowTop, visRight, rowTop)
        FormatLineShape shp, SHAPE_PREFIX & "RowLineTop", _
            modSettings.RowLineColor, modSettings.RowLineSize

        ' Bottom line
        Set shp = ws.Shapes.AddLine( _
            visLeft, rowBottom, visRight, rowBottom)
        FormatLineShape shp, SHAPE_PREFIX & "RowLineBottom", _
            modSettings.RowLineColor, modSettings.RowLineSize
    End If

    ' --- Col Lines (left + right) ---
    If modSettings.ColLineEnabled Then
        ' Left line
        Set shp = ws.Shapes.AddLine( _
            colLeft, visTop, colLeft, visBottom)
        FormatLineShape shp, SHAPE_PREFIX & "ColLineLeft", _
            modSettings.ColLineColor, modSettings.ColLineSize

        ' Right line
        Set shp = ws.Shapes.AddLine( _
            colRight, visTop, colRight, visBottom)
        FormatLineShape shp, SHAPE_PREFIX & "ColLineRight", _
            modSettings.ColLineColor, modSettings.ColLineSize
    End If

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    ' Silently fail - don't break user's workflow
    Application.ScreenUpdating = True
    Debug.Print "RH DrawHighlights Error: " & Err.Description
End Sub

' --- Format a line shape ---
Private Sub FormatLineShape(ByVal shp As Shape, ByVal shapeName As String, _
                            ByVal lineColor As Long, ByVal lineWeight As Double)
    With shp
        .Name = shapeName
        .Line.ForeColor.RGB = lineColor
        .Line.Weight = lineWeight
        .Line.Visible = msoTrue
        .Placement = xlFreeFloating
    End With
End Sub

' --- Quick toggles (callable from shortcuts) ---
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
