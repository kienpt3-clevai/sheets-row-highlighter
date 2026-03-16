(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory()
    return
  }

  root.PopupSettingsUtils = factory()
})(typeof globalThis !== 'undefined' ? globalThis : this, () => {
  /**
   * Chuẩn hoá settings dùng cho Reset:
   * - Ưu tiên defaultSettings người dùng đã lưu
   * - Fallback về DEFAULT_* hard-code khi chưa có hoặc thiếu field
   *
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
  const getResetSettings = (storedDefaults, fallbackDefaults) => {
    const current = storedDefaults || {}
    const baseColor = current.color ?? fallbackDefaults.defaultColor

    return {
      color: baseColor,
      rowLineColor: current.rowLineColor ?? baseColor ?? fallbackDefaults.defaultRowLineColor,
      colLineColor: current.colLineColor ?? baseColor ?? fallbackDefaults.defaultColLineColor,
      rowFillColor: current.rowFillColor ?? baseColor ?? fallbackDefaults.defaultRowFillColor,
      colFillColor: current.colFillColor ?? baseColor ?? fallbackDefaults.defaultColFillColor,
      opacity: current.opacity ?? fallbackDefaults.defaultOpacity,
      row: current.row ?? fallbackDefaults.defaultRow,
      column: current.column ?? fallbackDefaults.defaultColumn,
      fillRow: current.fillRow ?? fallbackDefaults.defaultFillRow,
      fillCol: current.fillCol ?? fallbackDefaults.defaultFillCol,
      lineSize: current.lineSize ?? fallbackDefaults.defaultLineSize,
      rowFillOpacity:
        current.rowFillOpacity ?? current.cellOpacity ?? fallbackDefaults.defaultRowFillOpacity,
      colFillOpacity:
        current.colFillOpacity ?? current.cellOpacity ?? fallbackDefaults.defaultColFillOpacity,
    }
  }

  return { getResetSettings }
})
