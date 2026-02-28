// @ts-check
/**
 * Zoom in/out trên Google Sheets: mô phỏng Alt+/, gõ "Zoom: X%", Enter.
 * Đọc zoom hiện tại từ DOM toolbar (aria-label "Zoom list. X% selected.").
 */

const ZOOM_LEVELS = [50, 75, 90, 100, 125, 150, 200]

const ZOOM_MENU_SEARCH_DELAY_MS = 40   // Chờ menu mở
const ZOOM_ENTER_DELAY_MS = 130        // Chờ menu lọc xong rồi gửi Enter
const ZOOM_AFTER_ENTER_MS = 30         // Chờ đóng menu

const MENU_SEARCH_SELECTORS = [
  'input.docs-omnibox-input',
  'input[aria-label="Menus"]',
  'input[placeholder="Menus"]',
  'input[aria-label*="Search"]',
  'input[placeholder*="Search"]',
  'input[placeholder*="earch"]',
  'input.jfk-textinput',
  'input[type="text"]',
]

/**
 * Tìm ô nhập văn bản hiển thị trong node và shadow roots / iframes.
 * @param {Document | DocumentFragment | Element} node
 * @returns {HTMLInputElement | null}
 */
function findMenuSearchInputInDoc(node) {
  if (!node || typeof node.querySelectorAll !== 'function') return null
  const root = node instanceof Document ? node : node.getRootNode ? node.getRootNode() : node
  const doc = root instanceof Document ? root : null
  if (!doc) return null
  for (const sel of MENU_SEARCH_SELECTORS) {
    const list = doc.querySelectorAll(sel)
    for (let i = 0; i < list.length; i++) {
      const el = list[i]
      if (el && el instanceof HTMLInputElement && el.offsetParent !== null) return el
    }
  }
  const walk = (/** @type {Element} */ el) => {
    if (el.shadowRoot) {
      const inShadow = findMenuSearchInputInDoc(el.shadowRoot)
      if (inShadow) return inShadow
    }
    for (const child of el.children) {
      const found = walk(child)
      if (found) return found
    }
    return null
  }
  const body = doc.body || doc.documentElement
  if (body) {
    const inShadow = walk(body)
    if (inShadow) return inShadow
  }
  try {
    const frames = doc.querySelectorAll('iframe')
    for (const fr of frames) {
      try {
        const sub = fr.contentDocument
        if (sub) {
          const found = findMenuSearchInputInDoc(sub)
          if (found) return found
        }
      } catch (_) {}
    }
  } catch (_) {}
  return null
}

/**
 * Tìm ô "Search the Menus" (doc hiện tại, doc top, hoặc iframe).
 * @param {Document} doc
 * @returns {HTMLInputElement | null}
 */
function findMenuSearchInput(doc) {
  let el = findMenuSearchInputInDoc(doc)
  if (el) return el
  try {
    if (window.top && window.top.document && window.top.document !== doc) {
      el = findMenuSearchInputInDoc(window.top.document)
      if (el) return el
    }
  } catch (_) {}
  return null
}

/**
 * @param {Document} doc
 * @returns {number} Phần trăm zoom hiện tại, hoặc 100 nếu không đọc được.
 */
function getCurrentZoomPercent(doc) {
  const el = doc.querySelector('[role="option"][aria-label*="Zoom list"]')
  if (!el || typeof el.getAttribute !== 'function') return 100
  const label = el.getAttribute('aria-label') || ''
  const match = label.match(/(\d+)\s*%\s*selected/i)
  const percent = match ? parseInt(match[1], 10) : 100
  if (!Number.isFinite(percent) || percent <= 0) return 100
  return percent
}

/**
 * Mức zoom đích sau khi dịch N bước (dương = phóng to, âm = thu nhỏ).
 * @param {number} current
 * @param {number} steps
 * @returns {number} Phần trăm đích trong ZOOM_LEVELS.
 */
function getTargetZoomPercentBySteps(current, steps) {
  const idx = ZOOM_LEVELS.indexOf(current)
  const i = idx < 0 ? ZOOM_LEVELS.findIndex((p) => p >= current) : idx
  const baseIdx = i < 0 ? 0 : i
  const newIdx = Math.max(0, Math.min(ZOOM_LEVELS.length - 1, baseIdx + steps))
  return ZOOM_LEVELS[newIdx]
}

/** @type {Record<string, string>} */
const KEY_TO_CODE = {
  ' ': 'Space',
  '%': 'Digit5',
  ':': 'Semicolon',
}

/**
 * Gửi sự kiện phím lên target.
 * @param {Element} target
 * @param {string} type
 * @param {string} key
 * @param {string} code
 * @param {boolean} [altKey]
 */
function dispatchKey(target, type, key, code, altKey = false) {
  const keyCode = key === 'Enter' ? 13 : key.length === 1 ? key.charCodeAt(0) : 0
  const opts = {
    key,
    code: code || KEY_TO_CODE[key] || (key.length === 1 ? `Key${key.toUpperCase()}` : ''),
    keyCode,
    which: keyCode,
    altKey: !!altKey,
    bubbles: true,
    cancelable: true,
  }
  target.dispatchEvent(new KeyboardEvent(type, opts))
}

/**
 * Mở Search the Menus (Alt+/), gõ "Zoom: X%" rồi Enter.
 * @param {Document} doc
 * @param {number} targetPercent
 * @param {() => void} [onComplete] Gọi sau khi gửi Enter và đợi đóng menu.
 */
function applyZoom(doc, targetPercent, onComplete) {
  const root = doc.documentElement || doc.body

  dispatchKey(root, 'keydown', '/', 'Slash', true)
  dispatchKey(root, 'keyup', '/', 'Slash', true)

  function fillSearchAndEnter() {
    const searchText = `Zoom: ${targetPercent}%`
    const searchInput = findMenuSearchInput(doc)
    if (searchInput) {
      try {
        searchInput.focus()
        searchInput.select()
        searchInput.value = searchText
        searchInput.dispatchEvent(new Event('input', { bubbles: true }))
        const sendEnter = () => {
          const target = doc.activeElement || searchInput
          const opts = { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true }
          target.dispatchEvent(new KeyboardEvent('keydown', opts))
          target.dispatchEvent(new KeyboardEvent('keypress', opts))
          target.dispatchEvent(new KeyboardEvent('keyup', opts))
          if (typeof onComplete === 'function') {
            setTimeout(onComplete, ZOOM_AFTER_ENTER_MS)
          }
        }
        setTimeout(sendEnter, ZOOM_ENTER_DELAY_MS)
      } catch (_) {
        if (typeof onComplete === 'function') onComplete()
      }
      return true
    }
    return false
  }

  function fallbackKeyDispatch() {
    const searchText = `Zoom: ${targetPercent}%`
    let target = doc.activeElement
    if (!target || target === doc.body) target = root
    for (const char of searchText) {
      const code = KEY_TO_CODE[char] || (char.length === 1 ? `Key${char.toUpperCase()}` : '')
      dispatchKey(target, 'keydown', char, code)
      dispatchKey(target, 'keypress', char, code)
      dispatchKey(target, 'keyup', char, code)
    }
    dispatchKey(target, 'keydown', 'Enter', 'Enter')
    dispatchKey(target, 'keyup', 'Enter', 'Enter')
    if (typeof onComplete === 'function') setTimeout(onComplete, ZOOM_AFTER_ENTER_MS)
  }

  setTimeout(() => {
    if (fillSearchAndEnter()) return
    fallbackKeyDispatch()
  }, ZOOM_MENU_SEARCH_DELAY_MS)
}

/** @type {number} */
let pendingZoomOut = 0
/** @type {number} */
let pendingZoomIn = 0
/** @type {ReturnType<typeof setTimeout> | null} */
let zoomDebounceTimer = null

const ZOOM_DEBOUNCE_MS = 200  // Chờ user thôi bấm rồi mới zoom một lần

/**
 * Áp dụng zoom đang chờ sau khi user thôi bấm (một lần tới mức đích).
 * @param {Document} doc
 */
function applyPendingZoom(doc) {
  zoomDebounceTimer = null
  if (pendingZoomOut > 0) {
    const steps = pendingZoomOut
    pendingZoomOut = 0
    pendingZoomIn = 0
    const current = getCurrentZoomPercent(doc)
    const target = getTargetZoomPercentBySteps(current, -steps)
    applyZoom(doc, target)
    return
  }
  if (pendingZoomIn > 0) {
    const steps = pendingZoomIn
    pendingZoomIn = 0
    pendingZoomOut = 0
    const current = getCurrentZoomPercent(doc)
    const target = getTargetZoomPercentBySteps(current, steps)
    applyZoom(doc, target)
  }
}

function scheduleApplyZoom(doc) {
  if (zoomDebounceTimer) clearTimeout(zoomDebounceTimer)
  zoomDebounceTimer = setTimeout(() => applyPendingZoom(doc), ZOOM_DEBOUNCE_MS)
}

/**
 * @param {Document} doc
 */
function zoomOut(doc) {
  pendingZoomOut++
  pendingZoomIn = 0
  scheduleApplyZoom(doc)
}

/**
 * @param {Document} doc
 */
function zoomIn(doc) {
  pendingZoomIn++
  pendingZoomOut = 0
  scheduleApplyZoom(doc)
}

// Cho main.js gọi (cùng global trong content script)
window.__SheetsZoom = { zoomIn, zoomOut }
