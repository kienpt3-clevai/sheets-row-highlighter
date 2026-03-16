const sendMessageToActiveTab = async (message) => {
  const [tab] = await chrome.tabs.query({
    active: true,
    lastFocusedWindow: true,
  })

  return chrome.tabs.sendMessage(tab.id, message)
}

window.addEventListener('load', () => {
  const rowInput = document.getElementById('row')
  const columnInput = document.getElementById('column')
  const rowLineSizeInput = document.getElementById('rowLineSize')
  const colLineSizeInput = document.getElementById('colLineSize')
  const rowOpacityInput = document.getElementById('rowOpacity')
  const colOpacityInput = document.getElementById('colOpacity')
  const fillRowInput = document.getElementById('fillRow')
  const fillColInput = document.getElementById('fillCol')
  const resetButton = document.getElementById('reset')
  const setDefaultButton = document.getElementById('setDefault')
  const setDefaultAllButton = document.getElementById('setDefaultAll')

  const defaultColor = typeof DEFAULT_COLOR !== 'undefined' ? DEFAULT_COLOR : '#c2185b'
  const defaultOpacity = typeof DEFAULT_OPACITY !== 'undefined' ? DEFAULT_OPACITY : '1'
  const defaultRow = typeof DEFAULT_ROW !== 'undefined' ? DEFAULT_ROW : true
  const defaultColumn = typeof DEFAULT_COLUMN !== 'undefined' ? DEFAULT_COLUMN : true
  const defaultFillRow = typeof DEFAULT_FILL_ROW !== 'undefined' ? DEFAULT_FILL_ROW : true
  const defaultFillCol = typeof DEFAULT_FILL_COL !== 'undefined' ? DEFAULT_FILL_COL : true
  const defaultLineSize = typeof DEFAULT_LINE_SIZE !== 'undefined' ? DEFAULT_LINE_SIZE : 3.25
  const defaultRowLineSize =
    typeof DEFAULT_ROW_LINE_SIZE !== 'undefined' ? DEFAULT_ROW_LINE_SIZE : defaultLineSize
  const defaultColLineSize =
    typeof DEFAULT_COL_LINE_SIZE !== 'undefined' ? DEFAULT_COL_LINE_SIZE : defaultLineSize
  const defaultRowFillOpacity =
    typeof DEFAULT_ROW_FILL_OPACITY !== 'undefined' ? DEFAULT_ROW_FILL_OPACITY : 0.05
  const defaultColFillOpacity =
    typeof DEFAULT_COL_FILL_OPACITY !== 'undefined' ? DEFAULT_COL_FILL_OPACITY : 0.05
  const defaultRowLineColor =
    typeof DEFAULT_ROW_LINE_COLOR !== 'undefined' ? DEFAULT_ROW_LINE_COLOR : defaultColor
  const defaultColLineColor =
    typeof DEFAULT_COL_LINE_COLOR !== 'undefined' ? DEFAULT_COL_LINE_COLOR : defaultColor
  const defaultRowFillColor =
    typeof DEFAULT_ROW_FILL_COLOR !== 'undefined' ? DEFAULT_ROW_FILL_COLOR : defaultColor
  const defaultColFillColor =
    typeof DEFAULT_COL_FILL_COLOR !== 'undefined' ? DEFAULT_COL_FILL_COLOR : defaultColor
  const fallbackDefaults = {
    defaultColor,
    defaultOpacity,
    defaultRow,
    defaultColumn,
    defaultFillRow,
    defaultFillCol,
    defaultRowLineSize,
    defaultColLineSize,
    defaultRowFillOpacity,
    defaultColFillOpacity,
    defaultRowLineColor,
    defaultColLineColor,
    defaultRowFillColor,
    defaultColFillColor,
  }
  const normalizePopupSettings =
    typeof PopupSettingsUtils !== 'undefined' &&
    PopupSettingsUtils &&
    typeof PopupSettingsUtils.normalizePopupSettings === 'function'
      ? PopupSettingsUtils.normalizePopupSettings
      : (sheetSettings, storedDefaults, fallback) => ({
          color: sheetSettings?.color ?? storedDefaults?.color ?? fallback.defaultColor,
          rowLineColor:
            sheetSettings?.rowLineColor ??
            storedDefaults?.rowLineColor ??
            sheetSettings?.color ??
            storedDefaults?.color ??
            fallback.defaultRowLineColor,
          colLineColor:
            sheetSettings?.colLineColor ??
            storedDefaults?.colLineColor ??
            sheetSettings?.color ??
            storedDefaults?.color ??
            fallback.defaultColLineColor,
          rowFillColor:
            sheetSettings?.rowFillColor ??
            storedDefaults?.rowFillColor ??
            sheetSettings?.color ??
            storedDefaults?.color ??
            fallback.defaultRowFillColor,
          colFillColor:
            sheetSettings?.colFillColor ??
            storedDefaults?.colFillColor ??
            sheetSettings?.color ??
            storedDefaults?.color ??
            fallback.defaultColFillColor,
          opacity: String(fallback.defaultOpacity),
          row: sheetSettings?.row ?? storedDefaults?.row ?? fallback.defaultRow,
          column: sheetSettings?.column ?? storedDefaults?.column ?? fallback.defaultColumn,
          fillRow: sheetSettings?.fillRow ?? storedDefaults?.fillRow ?? fallback.defaultFillRow,
          fillCol: sheetSettings?.fillCol ?? storedDefaults?.fillCol ?? fallback.defaultFillCol,
          rowLineSize:
            sheetSettings?.rowLineSize ??
            sheetSettings?.lineSize ??
            storedDefaults?.rowLineSize ??
            storedDefaults?.lineSize ??
            fallback.defaultRowLineSize,
          colLineSize:
            sheetSettings?.colLineSize ??
            sheetSettings?.lineSize ??
            storedDefaults?.colLineSize ??
            storedDefaults?.lineSize ??
            fallback.defaultColLineSize,
          rowFillOpacity:
            sheetSettings?.rowFillOpacity ??
            sheetSettings?.cellOpacity ??
            storedDefaults?.rowFillOpacity ??
            storedDefaults?.cellOpacity ??
            fallback.defaultRowFillOpacity,
          colFillOpacity:
            sheetSettings?.colFillOpacity ??
            sheetSettings?.cellOpacity ??
            storedDefaults?.colFillOpacity ??
            storedDefaults?.cellOpacity ??
            fallback.defaultColFillOpacity,
        })

  const customColors = [
    '#0e65eb',
    '#1565c0',
    '#1b5e20',
    '#5c0eec',
    '#6a1b9a',
    '#c2185b',
    '#ec0e2f',
    '#ec930e',
  ].sort()

  const huebeeOptions = {
    notation: 'hex',
    customColors,
    shades: 0,
    hues: 4,
  }

  const huebRowLine = new Huebee('#rowLineColor', huebeeOptions)
  const huebColLine = new Huebee('#colLineColor', huebeeOptions)
  const huebRowFill = new Huebee('#rowFillColor', huebeeOptions)
  const huebColFill = new Huebee('#colFillColor', huebeeOptions)

  const huebees = [
    { key: 'rowLineColor', hueb: huebRowLine },
    { key: 'colLineColor', hueb: huebColLine },
    { key: 'rowFillColor', hueb: huebRowFill },
    { key: 'colFillColor', hueb: huebColFill },
  ]

  const cycleColor = (hueb, delta) => {
    const current = hueb.color?.toLowerCase() ?? ''
    const idx = customColors.findIndex((c) => c.toLowerCase() === current)
    const len = customColors.length
    const nextIdx = idx < 0 ? 0 : (idx + delta + len) % len
    hueb.setColor(customColors[nextIdx])
    save()
  }

  const toFillOpacity = (input, fallback) =>
    Math.min(
      Math.max(
        Math.round((parseFloat(input.value, 10) || fallback) * 20) / 20,
        0
      ),
      0.5
    )

  /** Trả về object cấu hình từ giá trị hiện tại trên popup (dùng cho save / Set Default / Set Default All). */
  const getCurrentSettings = () => {
    const rowLineColor = huebRowLine.color ?? defaultRowLineColor
    const colLineColor = huebColLine.color ?? defaultColLineColor
    const rowFillColor = huebRowFill.color ?? defaultRowFillColor
    const colFillColor = huebColFill.color ?? defaultColFillColor
    const rowLineSize = Math.min(
      Math.max(parseFloat(rowLineSizeInput.value, 10) || defaultRowLineSize, 0.5),
      5
    )
    const colLineSize = Math.min(
      Math.max(parseFloat(colLineSizeInput.value, 10) || defaultColLineSize, 0.5),
      5
    )
    const rowFillOpacity = toFillOpacity(rowOpacityInput, defaultRowFillOpacity)
    const colFillOpacity = toFillOpacity(colOpacityInput, defaultColFillOpacity)
    return {
      color: rowLineColor,
      rowLineColor,
      colLineColor,
      rowFillColor,
      colFillColor,
      opacity: String(defaultOpacity),
      row: rowInput.checked,
      column: columnInput.checked,
      fillRow: fillRowInput.checked,
      fillCol: fillColInput.checked,
      rowLineSize,
      colLineSize,
      rowFillOpacity,
      colFillOpacity,
      cellOpacity: rowFillOpacity,
      updatedAt: Date.now(),
    }
  }

  const save = async () => {
    let sheetKey = 'default'
    try {
      const response = await sendMessageToActiveTab({ type: 'getSheetKey' })
      if (response && typeof response.sheetKey === 'string') {
        sheetKey = response.sheetKey
      }
    } catch {
      // ignore
    }

    chrome.storage.local.get(['sheetSettings'], (items) => {
      const allSettings = items.sheetSettings || {}
      const settings = getCurrentSettings()
      allSettings[sheetKey] = settings

      chrome.storage.local.set({ sheetSettings: allSettings }, () => {
        sendMessageToActiveTab({ type: 'settingsUpdated', settings, sheetKey }).catch(
          () => {}
        )
      })
    })
  }

  resetButton.addEventListener('click', () => {
    chrome.storage.local.get(['defaultSettings'], (items) => {
      const resetSettings = normalizePopupSettings(undefined, items.defaultSettings, fallbackDefaults)

      huebees.forEach(({ hueb }) => hueb.off('change', save))

      huebRowLine.setColor(resetSettings.rowLineColor)
      huebColLine.setColor(resetSettings.colLineColor)
      huebRowFill.setColor(resetSettings.rowFillColor)
      huebColFill.setColor(resetSettings.colFillColor)
      // Line opacity is fixed to 1.0 and no longer configurable.
      rowInput.checked = resetSettings.row
      columnInput.checked = resetSettings.column
      fillRowInput.checked = resetSettings.fillRow
      fillColInput.checked = resetSettings.fillCol
      rowLineSizeInput.value = resetSettings.rowLineSize
      colLineSizeInput.value = resetSettings.colLineSize
      rowOpacityInput.value = String(resetSettings.rowFillOpacity)
      colOpacityInput.value = String(resetSettings.colFillOpacity)

      void save()
      huebees.forEach(({ hueb }) => hueb.on('change', save))
    })
  })

  setDefaultButton.addEventListener('click', () => {
    chrome.storage.local.set({ defaultSettings: getCurrentSettings() })
  })

  setDefaultAllButton.addEventListener('click', () => {
    chrome.storage.local.get(['sheetSettings'], (items) => {
      const allSettings = items.sheetSettings || {}
      const settings = getCurrentSettings()
      for (const key of Object.keys(allSettings)) {
        allSettings[key] = { ...settings }
      }
      chrome.storage.local.set({ sheetSettings: allSettings }, () => {
        sendMessageToActiveTab({ type: 'settingsUpdated', settings }).catch(() => {})
      })
    })
  })

  document.getElementById('prevRowLineColor').addEventListener('click', () => cycleColor(huebRowLine, -1))
  document.getElementById('nextRowLineColor').addEventListener('click', () => cycleColor(huebRowLine, 1))
  document.getElementById('prevColLineColor').addEventListener('click', () => cycleColor(huebColLine, -1))
  document.getElementById('nextColLineColor').addEventListener('click', () => cycleColor(huebColLine, 1))
  document.getElementById('prevRowFillColor').addEventListener('click', () => cycleColor(huebRowFill, -1))
  document.getElementById('nextRowFillColor').addEventListener('click', () => cycleColor(huebRowFill, 1))
  document.getElementById('prevColFillColor').addEventListener('click', () => cycleColor(huebColFill, -1))
  document.getElementById('nextColFillColor').addEventListener('click', () => cycleColor(huebColFill, 1))

  ;(async () => {
    let sheetKey = 'default'
    try {
      const response = await sendMessageToActiveTab({ type: 'getSheetKey' })
      if (response && typeof response.sheetKey === 'string') {
        sheetKey = response.sheetKey
      }
    } catch {
      // ignore
    }

    chrome.storage.local.get(['sheetSettings', 'defaultSettings'], (items) => {
      const allSettings = items.sheetSettings || {}
      const defaultSettings = items.defaultSettings || {}
      const current = normalizePopupSettings(
        allSettings[sheetKey],
        defaultSettings,
        fallbackDefaults
      )

      huebRowLine.setColor(current.rowLineColor)
      huebColLine.setColor(current.colLineColor)
      huebRowFill.setColor(current.rowFillColor)
      huebColFill.setColor(current.colFillColor)
      // Line opacity is fixed to 1.0 and no longer configurable.
      rowInput.checked = current.row
      columnInput.checked = current.column
      fillRowInput.checked = current.fillRow
      fillColInput.checked = current.fillCol
      rowLineSizeInput.value = current.rowLineSize
      colLineSizeInput.value = current.colLineSize
      rowOpacityInput.value = String(current.rowFillOpacity)
      colOpacityInput.value = String(current.colFillOpacity)

      huebees.forEach(({ hueb }) => hueb.on('change', save))
      rowInput.addEventListener('change', () => void save())
      columnInput.addEventListener('change', () => void save())
      fillRowInput.addEventListener('change', () => void save())
      fillColInput.addEventListener('change', () => void save())
      rowLineSizeInput.addEventListener('change', () => void save())
      colLineSizeInput.addEventListener('change', () => void save())
      rowOpacityInput.addEventListener('change', () => void save())
      colOpacityInput.addEventListener('change', () => void save())
    })
  })()

  chrome.runtime.onMessage.addListener((request) => {
    if (request.type !== 'commands') return
    rowInput.checked = request.row
    columnInput.checked = request.column
  })
})
