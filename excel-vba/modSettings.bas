Attribute VB_Name = "modSettings"
Option Explicit

' === Registry constants ===
Public Const APP_NAME As String = "ExcelRowHighlighter"
Public Const SEC_GENERAL As String = "General"
Public Const SEC_DEFAULTS As String = "CustomDefaults"

' === Hard-coded default values ===
Public Const DEF_ROW_LINE_ENABLED As Boolean = True
Public Const DEF_COL_LINE_ENABLED As Boolean = True
Public Const DEF_ROW_FILL_ENABLED As Boolean = True
Public Const DEF_COL_FILL_ENABLED As Boolean = True
Public Const DEF_ROW_LINE_COLOR As String = "#c2185b"
Public Const DEF_COL_LINE_COLOR As String = "#3399ff"
Public Const DEF_ROW_FILL_COLOR As String = "#c2185b"
Public Const DEF_COL_FILL_COLOR As String = "#3399ff"
Public Const DEF_ROW_LINE_SIZE As Double = 2.25
Public Const DEF_COL_LINE_SIZE As Double = 1.5
Public Const DEF_ROW_FILL_OPACITY As Double = 0.15
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

' === Event handler (shared between ThisWorkbook and ShowSettings) ===
Public gAppEvents As clsAppEvents

' --- Hex to OLE Color conversion ---
Public Function HexToOLE(ByVal hex As String) As Long
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
Private Function RegGet(ByVal section As String, ByVal key As String, ByVal def As String) As String
    On Error Resume Next
    RegGet = GetSetting(APP_NAME, section, key, def)
End Function

' --- Write a string to Registry ---
Private Sub RegSet(ByVal section As String, ByVal key As String, ByVal val As String)
    SaveSetting APP_NAME, section, key, val
End Sub

' --- Load all settings from Registry into public vars ---
Public Sub LoadSettings()
    RowLineEnabled = CBool(RegGet(SEC_GENERAL, "RowLineEnabled", CStr(DEF_ROW_LINE_ENABLED)))
    ColLineEnabled = CBool(RegGet(SEC_GENERAL, "ColLineEnabled", CStr(DEF_COL_LINE_ENABLED)))
    RowFillEnabled = CBool(RegGet(SEC_GENERAL, "RowFillEnabled", CStr(DEF_ROW_FILL_ENABLED)))
    ColFillEnabled = CBool(RegGet(SEC_GENERAL, "ColFillEnabled", CStr(DEF_COL_FILL_ENABLED)))
    RowLineColor = HexToOLE(RegGet(SEC_GENERAL, "RowLineColor", DEF_ROW_LINE_COLOR))
    ColLineColor = HexToOLE(RegGet(SEC_GENERAL, "ColLineColor", DEF_COL_LINE_COLOR))
    RowFillColor = HexToOLE(RegGet(SEC_GENERAL, "RowFillColor", DEF_ROW_FILL_COLOR))
    ColFillColor = HexToOLE(RegGet(SEC_GENERAL, "ColFillColor", DEF_COL_FILL_COLOR))
    RowLineSize = CDbl(RegGet(SEC_GENERAL, "RowLineSize", CStr(DEF_ROW_LINE_SIZE)))
    ColLineSize = CDbl(RegGet(SEC_GENERAL, "ColLineSize", CStr(DEF_COL_LINE_SIZE)))
    RowFillOpacity = CDbl(RegGet(SEC_GENERAL, "RowFillOpacity", CStr(DEF_ROW_FILL_OPACITY)))
    ColFillOpacity = CDbl(RegGet(SEC_GENERAL, "ColFillOpacity", CStr(DEF_COL_FILL_OPACITY)))
End Sub

' --- Save all current settings to Registry ---
Public Sub SaveSettings()
    RegSet SEC_GENERAL, "RowLineEnabled", CStr(RowLineEnabled)
    RegSet SEC_GENERAL, "ColLineEnabled", CStr(ColLineEnabled)
    RegSet SEC_GENERAL, "RowFillEnabled", CStr(RowFillEnabled)
    RegSet SEC_GENERAL, "ColFillEnabled", CStr(ColFillEnabled)
    RegSet SEC_GENERAL, "RowLineColor", OLEToHex(RowLineColor)
    RegSet SEC_GENERAL, "ColLineColor", OLEToHex(ColLineColor)
    RegSet SEC_GENERAL, "RowFillColor", OLEToHex(RowFillColor)
    RegSet SEC_GENERAL, "ColFillColor", OLEToHex(ColFillColor)
    RegSet SEC_GENERAL, "RowLineSize", CStr(RowLineSize)
    RegSet SEC_GENERAL, "ColLineSize", CStr(ColLineSize)
    RegSet SEC_GENERAL, "RowFillOpacity", CStr(RowFillOpacity)
    RegSet SEC_GENERAL, "ColFillOpacity", CStr(ColFillOpacity)
End Sub

' --- Save current settings as custom defaults ---
Public Sub SaveCustomDefaults()
    RegSet SEC_DEFAULTS, "RowLineEnabled", CStr(RowLineEnabled)
    RegSet SEC_DEFAULTS, "ColLineEnabled", CStr(ColLineEnabled)
    RegSet SEC_DEFAULTS, "RowFillEnabled", CStr(RowFillEnabled)
    RegSet SEC_DEFAULTS, "ColFillEnabled", CStr(ColFillEnabled)
    RegSet SEC_DEFAULTS, "RowLineColor", OLEToHex(RowLineColor)
    RegSet SEC_DEFAULTS, "ColLineColor", OLEToHex(ColLineColor)
    RegSet SEC_DEFAULTS, "RowFillColor", OLEToHex(RowFillColor)
    RegSet SEC_DEFAULTS, "ColFillColor", OLEToHex(ColFillColor)
    RegSet SEC_DEFAULTS, "RowLineSize", CStr(RowLineSize)
    RegSet SEC_DEFAULTS, "ColLineSize", CStr(ColLineSize)
    RegSet SEC_DEFAULTS, "RowFillOpacity", CStr(RowFillOpacity)
    RegSet SEC_DEFAULTS, "ColFillOpacity", CStr(ColFillOpacity)
    RegSet SEC_DEFAULTS, "HasCustom", "True"
End Sub

' --- Reset to custom defaults (or hard-coded if no custom saved) ---
Public Sub ResetSettings()
    Dim hasCustom As Boolean
    hasCustom = CBool(RegGet(SEC_DEFAULTS, "HasCustom", "False"))

    If hasCustom Then
        ' Load from custom defaults
        RowLineEnabled = CBool(RegGet(SEC_DEFAULTS, "RowLineEnabled", CStr(DEF_ROW_LINE_ENABLED)))
        ColLineEnabled = CBool(RegGet(SEC_DEFAULTS, "ColLineEnabled", CStr(DEF_COL_LINE_ENABLED)))
        RowFillEnabled = CBool(RegGet(SEC_DEFAULTS, "RowFillEnabled", CStr(DEF_ROW_FILL_ENABLED)))
        ColFillEnabled = CBool(RegGet(SEC_DEFAULTS, "ColFillEnabled", CStr(DEF_COL_FILL_ENABLED)))
        RowLineColor = HexToOLE(RegGet(SEC_DEFAULTS, "RowLineColor", DEF_ROW_LINE_COLOR))
        ColLineColor = HexToOLE(RegGet(SEC_DEFAULTS, "ColLineColor", DEF_COL_LINE_COLOR))
        RowFillColor = HexToOLE(RegGet(SEC_DEFAULTS, "RowFillColor", DEF_ROW_FILL_COLOR))
        ColFillColor = HexToOLE(RegGet(SEC_DEFAULTS, "ColFillColor", DEF_COL_FILL_COLOR))
        RowLineSize = CDbl(RegGet(SEC_DEFAULTS, "RowLineSize", CStr(DEF_ROW_LINE_SIZE)))
        ColLineSize = CDbl(RegGet(SEC_DEFAULTS, "ColLineSize", CStr(DEF_COL_LINE_SIZE)))
        RowFillOpacity = CDbl(RegGet(SEC_DEFAULTS, "RowFillOpacity", CStr(DEF_ROW_FILL_OPACITY)))
        ColFillOpacity = CDbl(RegGet(SEC_DEFAULTS, "ColFillOpacity", CStr(DEF_COL_FILL_OPACITY)))
    Else
        ' Use hard-coded defaults
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
    End If
    SaveSettings
End Sub

' --- Show the settings UserForm ---
Public Sub ShowSettings()
    frmSettings.Show vbModal
    ' Re-initialize event handler after modal form closes
    InitEventHandler
End Sub

' --- (Re)initialize the application event handler ---
Public Sub InitEventHandler()
    If gAppEvents Is Nothing Then
        Set gAppEvents = New clsAppEvents
    End If
    Set gAppEvents.App = Application
End Sub
