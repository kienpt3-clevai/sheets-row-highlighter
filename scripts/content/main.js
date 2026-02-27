// @ts-check
/// <reference path="./SheetsActiveCellLocator.js" />
/// <reference path="./ExcelActiveCellLocator.js" />
/// <reference path="./RowHighlighterApp.js" />

const isSheetsHost = location.host === 'docs.google.com'

const appContainer = document.createElement('div')
appContainer.id = 'rh-app-container'
document.body.appendChild(appContainer)

const locator = isSheetsHost
  ? new SheetsActiveCellLocator()
  : new ExcelActiveCellLocator()

const app = new RowHighlighterApp(appContainer, locator)
const updateHighlight = app.update.bind(app)

window.addEventListener('click', updateHighlight)
window.addEventListener('keydown', updateHighlight)
window.addEventListener('keyup', updateHighlight)
window.addEventListener('resize', updateHighlight)
window.addEventListener('scroll', updateHighlight, true)

// @ts-ignore chrome.xxxの参照エラーを無視
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
    app.headerColScale = current.headerColScale ?? app.headerColScale
    app.headerRowScale = current.headerRowScale ?? app.headerRowScale

    updateHighlight()
  })
}
loadSettings()

storage.onChanged.addListener(loadSettings)

// 設定更新時の再描画 / sheetKey 取得
// @ts-ignore chrome.xxxの参照エラーを無視
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.type === 'settingsUpdated') {
    loadSettings()
  }
  if (message.type === 'getSheetKey') {
    const sheetKey =
      typeof locator.getSheetKey === 'function' ? locator.getSheetKey() : 'default'
    sendResponse({ sheetKey })
  }
  return true
})
