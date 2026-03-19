Attribute VB_Name = "modSettings"
Option Explicit

' === Registry constants ===
Public Const APP_NAME As String = "ExcelRowHighlighter"
Public Const SEC_GENERAL As String = "General"

' === Default values (matching Chrome extension) ===
Public Const DEF_ROW_LINE_ENABLED As Boolean = True
Public Const DEF_COL_LINE_ENABLED As Boolean = True
Public Const DEF_ROW_FILL_ENABLED As Boolean = True
Public Const DEF_COL_FILL_ENABLED As Boolean = True
Public Const DEF_ROW_LINE_COLOR As String = "#c2185b"
Public Const DEF_COL_LINE_COLOR As String = "#c2185b"
Public Const DEF_ROW_FILL_COLOR As String = "#c2185b"
Public Const DEF_COL_FILL_COLOR As String = "#c2185b"
Public Const DEF_ROW_LINE_SIZE As Double = 3.25
Public Const DEF_COL_LINE_SIZE As Double = 3.25
Public Const DEF_ROW_FILL_OPACITY As Double = 0.05
Public Const DEF_COL_FILL_OPACITY As Double = 0.05

' === Cached settings (loaded once, updated by UserForm) ===
Public RowLineEnabled As Boolean
Public ColLineEnabled As Boolean
Public RowFillEnabled As Boolean
Public ColFillEnabled As Boolean
Public RowLineColor As Long    ' OLE color
Public ColLineColor As Long
Public RowFillColor As Long
Public ColFillColor As Long
Public RowLineSize As Double
Public ColLineSize As Double
Public RowFillOpacity As Double
Public ColFillOpacity As Double

' --- Hex to OLE Color conversion ---
Public Function HexToOLE(ByVal hex As String) As Long
    ' Convert "#RRGGBB" to OLE color (Long = &HBBGGRR)
    If Left$(hex, 1) = "#" Then hex = Mid$(hex, 2)
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid$(hex, 1, 2))
    g = CLng("&H" & Mid$(hex, 3, 2))
    b = CLng("&H" & Mid$(hex, 5, 2))
    HexToOLE = RGB(r, g, b)
End Function

' --- OLE Color to Hex conversion ---
Public Function OLEToHex(ByVal oleColor As Long) As String
    Dim r As Long, g As Long, b As Long
    r = oleColor Mod 256
    g = (oleColor \ 256) Mod 256
    b = (oleColor \ 65536) Mod 256
    OLEToHex = "#" & LCase$(Right$("0" & Hex$(r), 2) & _
                             Right$("0" & Hex$(g), 2) & _
                             Right$("0" & Hex$(b), 2))
End Function

' --- Read a string from Registry with default ---
Private Function RegGet(ByVal key As String, ByVal def As String) As String
    On Error Resume Next
    RegGet = GetSetting(APP_NAME, SEC_GENERAL, key, def)
End Function

' --- Write a string to Registry ---
Private Sub RegSet(ByVal key As String, ByVal val As String)
    SaveSetting APP_NAME, SEC_GENERAL, key, val
End Sub

' --- Load all settings from Registry into public vars ---
Public Sub LoadSettings()
    RowLineEnabled = CBool(RegGet("RowLineEnabled", CStr(DEF_ROW_LINE_ENABLED)))
    ColLineEnabled = CBool(RegGet("ColLineEnabled", CStr(DEF_COL_LINE_ENABLED)))
    RowFillEnabled = CBool(RegGet("RowFillEnabled", CStr(DEF_ROW_FILL_ENABLED)))
    ColFillEnabled = CBool(RegGet("ColFillEnabled", CStr(DEF_COL_FILL_ENABLED)))
    RowLineColor = HexToOLE(RegGet("RowLineColor", DEF_ROW_LINE_COLOR))
    ColLineColor = HexToOLE(RegGet("ColLineColor", DEF_COL_LINE_COLOR))
    RowFillColor = HexToOLE(RegGet("RowFillColor", DEF_ROW_FILL_COLOR))
    ColFillColor = HexToOLE(RegGet("ColFillColor", DEF_COL_FILL_COLOR))
    RowLineSize = CDbl(RegGet("RowLineSize", CStr(DEF_ROW_LINE_SIZE)))
    ColLineSize = CDbl(RegGet("ColLineSize", CStr(DEF_COL_LINE_SIZE)))
    RowFillOpacity = CDbl(RegGet("RowFillOpacity", CStr(DEF_ROW_FILL_OPACITY)))
    ColFillOpacity = CDbl(RegGet("ColFillOpacity", CStr(DEF_COL_FILL_OPACITY)))
End Sub

' --- Save all settings to Registry from public vars ---
Public Sub SaveSettings()
    RegSet "RowLineEnabled", CStr(RowLineEnabled)
    RegSet "ColLineEnabled", CStr(ColLineEnabled)
    RegSet "RowFillEnabled", CStr(RowFillEnabled)
    RegSet "ColFillEnabled", CStr(ColFillEnabled)
    RegSet "RowLineColor", OLEToHex(RowLineColor)
    RegSet "ColLineColor", OLEToHex(ColLineColor)
    RegSet "RowFillColor", OLEToHex(RowFillColor)
    RegSet "ColFillColor", OLEToHex(ColFillColor)
    RegSet "RowLineSize", CStr(RowLineSize)
    RegSet "ColLineSize", CStr(ColLineSize)
    RegSet "RowFillOpacity", CStr(RowFillOpacity)
    RegSet "ColFillOpacity", CStr(ColFillOpacity)
End Sub

' --- Reset all settings to defaults ---
Public Sub ResetSettings()
    RowLineEnabled = DEF_ROW_LINE_ENABLED
    ColLineEnabled = DEF_COL_LINE_ENABLED
    RowFillEnabled = DEF_ROW_FILL_ENABLED
    ColFillEnabled = DEF_COL_FILL_ENABLED
    RowLineColor = HexToOLE(DEF_ROW_LINE_COLOR)
    ColLineColor = HexToOLE(DEF_COL_LINE_COLOR)
    RowFillColor = HexToOLE(DEF_ROW_FILL_COLOR)
    ColFillColor = HexToOLE(DEF_COL_FILL_COLOR)
    RowLineSize = DEF_ROW_LINE_SIZE
    ColLineSize = DEF_COL_LINE_SIZE
    RowFillOpacity = DEF_ROW_FILL_OPACITY
    ColFillOpacity = DEF_COL_FILL_OPACITY
    SaveSettings
End Sub

' --- Show the settings UserForm ---
Public Sub ShowSettings()
    frmSettings.Show vbModal
End Sub
