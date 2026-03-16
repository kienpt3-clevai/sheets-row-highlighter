(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory()
    return
  }

  root.PopupSettingsUtils = factory()
})(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  /**
   * Chuẩn hoá settings cho popup:
   * - sheetSettings hiện tại ưu tiên cao nhất
   * - sau đó tới defaultSettings do người dùng lưu
   * - cuối cùng fallback về DEFAULT_* hard-code
   *
   * @param {Record<string, any> | undefined} sheetSettings
   * @param {Record<string, any> | undefined} storedDefaults
   * @param {{
   *   defaultColor: string
   *   defaultOpacity: string | number
   *   defaultRow: boolean
   *   defaultColumn: boolean
   *   defaultFillRow: boolean
   *   defaultFillCol: boolean
   *   defaultLineSize: number
   *   defaultRowFillOpacity: number
   *   defaultColFillOpacity: number
   *   defaultRowLineColor: string
   *   defaultColLineColor: string
   *   defaultRowFillColor: string
   *   defaultColFillColor: string
   * }} fallbackDefaults
   */
  const normalizePopupSettings = (sheetSettings, storedDefaults, fallbackDefaults) => {
    const defaults = storedDefaults || {}
    const current = sheetSettings || {}
    const baseColor = current.color ?? defaults.color ?? fallbackDefaults.defaultColor

    return {
      color: baseColor,
      rowLineColor:
        current.rowLineColor ??
        defaults.rowLineColor ??
        baseColor ??
        fallbackDefaults.defaultRowLineColor,
      colLineColor:
        current.colLineColor ??
        defaults.colLineColor ??
        baseColor ??
        fallbackDefaults.defaultColLineColor,
      rowFillColor:
        current.rowFillColor ??
        defaults.rowFillColor ??
        baseColor ??
        fallbackDefaults.defaultRowFillColor,
      colFillColor:
        current.colFillColor ??
        defaults.colFillColor ??
        baseColor ??
        fallbackDefaults.defaultColFillColor,
      opacity: current.opacity ?? defaults.opacity ?? fallbackDefaults.defaultOpacity,
      row: current.row ?? defaults.row ?? fallbackDefaults.defaultRow,
      column: current.column ?? defaults.column ?? fallbackDefaults.defaultColumn,
      fillRow: current.fillRow ?? defaults.fillRow ?? fallbackDefaults.defaultFillRow,
      fillCol: current.fillCol ?? defaults.fillCol ?? fallbackDefaults.defaultFillCol,
      lineSize: current.lineSize ?? defaults.lineSize ?? fallbackDefaults.defaultLineSize,
      rowFillOpacity:
        current.rowFillOpacity ??
        current.cellOpacity ??
        defaults.rowFillOpacity ??
        defaults.cellOpacity ??
        fallbackDefaults.defaultRowFillOpacity,
      colFillOpacity:
        current.colFillOpacity ??
        current.cellOpacity ??
        defaults.colFillOpacity ??
        defaults.cellOpacity ??
        fallbackDefaults.defaultColFillOpacity,
    }
  }

  const getResetSettings = (storedDefaults, fallbackDefaults) =>
    normalizePopupSettings(undefined, storedDefaults, fallbackDefaults)

  return { normalizePopupSettings, getResetSettings }
})
