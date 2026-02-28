// @ts-check
/// <reference path="./SheetsActiveCellLocator.js" />
/// <reference path="./RowHighlighterApp.js" />

/** @type {any} */
const chrome = (/** @type {{ chrome?: any }} */ (globalThis)).chrome

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

const loadSettings = () => {
  const sheetKey =
    typeof locator.getSheetKey === 'function' ? locator.getSheetKey() : 'default'

  storage.local.get(['sheetSettings'], (/** @type {any} */ items) => {
    /** @type {Record<string, any>} */
    const allSettings = items.sheetSettings || {}

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

    const current = pruned[sheetKey] || {}

    app.backgroundColor = current.color ?? app.backgroundColor
    app.opacity = current.opacity ?? app.opacity
    app.isRowEnabled = current.row ?? app.isRowEnabled
    app.isColEnabled = current.column ?? app.isColEnabled
    app.lineSize = current.lineSize ?? app.lineSize
    app.cellOpacity = current.cellOpacity ?? app.cellOpacity

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
      const current = message.settings
      app.backgroundColor = current.color ?? app.backgroundColor
      app.opacity = current.opacity ?? app.opacity
      app.isRowEnabled = current.row ?? app.isRowEnabled
      app.isColEnabled = current.column ?? app.isColEnabled
      app.lineSize = current.lineSize ?? app.lineSize
      app.cellOpacity = current.cellOpacity ?? app.cellOpacity

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
