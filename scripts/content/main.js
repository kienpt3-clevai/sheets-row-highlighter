// @ts-check
/// <reference path="./global.d.ts" />
/// <reference path="./SheetsActiveCellLocator.js" />
/// <reference path="./RowHighlighterApp.js" />

const appContainer = document.createElement('div')
appContainer.id = 'rh-app-container'
document.body.appendChild(appContainer)

const locator = new SheetsActiveCellLocator()

const app = new RowHighlighterApp(appContainer, locator)
const updateHighlight = app.update.bind(app)

window.addEventListener('click', updateHighlight)
window.addEventListener('keydown', updateHighlight)
window.addEventListener('keyup', updateHighlight)
window.addEventListener('resize', updateHighlight)
window.addEventListener('scroll', updateHighlight, true)

// Phím tắt zoom: Ctrl+, thu nhỏ, Ctrl+. phóng to (chỉ frame chính)
const win = /** @type {Window & { __SheetsZoom?: { zoomIn: (d: Document) => void; zoomOut: (d: Document) => void } }} */ (window)
const sheetsZoom = win.__SheetsZoom
if (win === window.top && sheetsZoom) {
  document.addEventListener(
    'keydown',
    (e) => {
      if (!e.ctrlKey) return
      if (e.key === ',') {
        e.preventDefault()
        e.stopPropagation()
        sheetsZoom.zoomOut(document)
      } else if (e.key === '.') {
        e.preventDefault()
        e.stopPropagation()
        sheetsZoom.zoomIn(document)
      }
    },
    true
  )
}

const storage = chrome.storage
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

const applySettingsToApp = (settings) => {
  app.backgroundColor = settings.color
  app.rowLineColor = settings.rowLineColor
  app.colLineColor = settings.colLineColor
  app.rowFillColor = settings.rowFillColor
  app.colFillColor = settings.colFillColor
  app.opacity = settings.opacity
  app.isRowEnabled = settings.row
  app.isColEnabled = settings.column
  app.fillRowEnabled = settings.fillRow
  app.fillColEnabled = settings.fillCol
  app.rowLineSize = settings.rowLineSize
  app.colLineSize = settings.colLineSize
  app.rowFillOpacity = settings.rowFillOpacity
  app.colFillOpacity = settings.colFillOpacity
}

const loadSettings = () => {
  const sheetKey =
    typeof locator.getSheetKey === 'function' ? locator.getSheetKey() : 'default'

  storage.local.get(['sheetSettings', 'defaultSettings'], (/** @type {any} */ items) => {
    /** @type {Record<string, any>} */
    const allSettings = items.sheetSettings || {}
    /** @type {Record<string, any>} */
    const defaultSettings = items.defaultSettings || {}

    // Tự động xoá cấu hình cũ hơn 30 ngày
    const now = Date.now()
    const THIRTY_DAYS = 30 * 24 * 60 * 60 * 1000
    /** @type {Record<string, any>} */
    const pruned = {}
    Object.keys(allSettings).forEach((key) => {
      const value = allSettings[key]
      if (!value || typeof value !== 'object') {
        return
      }
      const updatedAt = typeof value.updatedAt === 'number' ? value.updatedAt : 0
      if (!updatedAt || now - updatedAt <= THIRTY_DAYS) {
        pruned[key] = value
      }
    })

    if (Object.keys(pruned).length !== Object.keys(allSettings).length) {
      storage.local.set({ sheetSettings: pruned })
    }

    const current = normalizePopupSettings(pruned[sheetKey], defaultSettings, fallbackDefaults)
    applySettingsToApp(current)

    updateHighlight()
  })
}
loadSettings()

storage.onChanged.addListener(loadSettings)

// Cập nhật khi popup/background gửi message; trả sheetKey khi được hỏi
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.type === 'applyCommand') {
    app.isRowEnabled = message.row ?? app.isRowEnabled
    app.isColEnabled = message.column ?? app.isColEnabled
    updateHighlight()
    // Đồng bộ vào sheetSettings để lần sau load đúng
    const sheetKey =
      typeof locator.getSheetKey === 'function' ? locator.getSheetKey() : 'default'
    chrome.storage.local.get(['sheetSettings'], (items) => {
      const all = items.sheetSettings || {}
      const cur = all[sheetKey] || {}
      all[sheetKey] = {
        ...cur,
        row: app.isRowEnabled,
        column: app.isColEnabled,
        updatedAt: Date.now(),
      }
      chrome.storage.local.set({ sheetSettings: all })
    })
  }
  if (message.type === 'settingsUpdated') {
    if (message.settings && typeof message.settings === 'object') {
      const current = normalizePopupSettings(message.settings, undefined, fallbackDefaults)
      applySettingsToApp(current)
      updateHighlight()
    } else {
      loadSettings()
    }
  }
  if (message.type === 'getSheetKey') {
    const sheetKey =
      typeof locator.getSheetKey === 'function' ? locator.getSheetKey() : 'default'
    sendResponse({ sheetKey })
  }
  return true
})
