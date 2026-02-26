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
  storage.local.get(
    ['color', 'opacity', 'row', 'column', 'lineSize'],
    (/** @type {any} */ items) => {
      app.backgroundColor = items.color ?? app.backgroundColor
      app.opacity = items.opacity ?? app.opacity
      app.isRowEnabled = items.row ?? app.isRowEnabled
      app.isColEnabled = items.column ?? app.isColEnabled
      app.lineSize = items.lineSize ?? app.lineSize
      updateHighlight()
    }
  )
}
loadSettings()

storage.onChanged.addListener(loadSettings)

/**
 * Apply Google Sheets toolbar zoom to the given percentage (e.g. 75, 100).
 * Only runs on docs.google.com; safely no-ops elsewhere.
 * @param {number} targetZoom
 */
const applySheetsZoom = (targetZoom) => {
  if (!isSheetsHost) return
  if (typeof targetZoom !== 'number') return

  try {
    const doc = document

    // Try to find the Zoom dropdown button on the toolbar
    const zoomButton =
      doc.querySelector('div[role="button"][aria-label*="Zoom"]') ||
      doc.querySelector('div[role="button"][data-tooltip*="Zoom"]')

    if (!zoomButton) return

    zoomButton.click()

    const label = `${targetZoom}%`

    const trySelect = () => {
      const menu =
        doc.querySelector('div[role="menu"]') ||
        doc.querySelector('[role="menu"][jsname]')
      if (!menu) return

      const items = Array.from(
        menu.querySelectorAll(
          '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]'
        )
      )

      const targetItem = items.find((el) => {
        const text = el.textContent
        return text && text.trim() === label
      })

      if (targetItem instanceof HTMLElement) {
        targetItem.click()
      }
    }

    // Allow menu DOM to render before querying items
    setTimeout(trySelect, 50)
  } catch {
    // Fail silently; do not break other features
  }
}

// 設定更新時の再描画 & Zoom コマンド
// @ts-ignore chrome.xxxの参照エラーを無視
chrome.runtime.onMessage.addListener((message) => {
  if (message.type === 'settingsUpdated') {
    loadSettings()
  } else if (message.type === 'zoomCommand') {
    applySheetsZoom(message.zoomLevel)
  }
})
