// @ts-check
/// <reference path="./SheetsActiveCellLocator.js" />
/// <reference path="./ExcelActiveCellLocator.js" />
/// <reference path="./RowHighlighterApp.js" />

const appContainer = document.createElement('div')
appContainer.id = 'rh-app-container'
document.body.appendChild(appContainer)

const locator =
  location.host === 'docs.google.com'
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

// 設定更新時の再描画
// @ts-ignore chrome.xxxの参照エラーを無視
chrome.runtime.onMessage.addListener((message) => {
  if (message.type === 'settingsUpdated') {
    loadSettings()
  }
})
