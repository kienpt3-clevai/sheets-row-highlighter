// @ts-check
// Google Slides: Ctrl+Alt+/ → gửi Ctrl+Alt+[ (fit to screen)

const FIT_KEY = '['
const FIT_CODE = 'BracketLeft'
const FIT_KEY_CODE = 219

function isSlashKey(e) {
  return e.key === '/' || e.key === 'Slash' || e.code === 'Slash' || e.keyCode === 191
}

function dispatchFitToScreen(target) {
  const opts = {
    key: FIT_KEY,
    code: FIT_CODE,
    keyCode: FIT_KEY_CODE,
    which: FIT_KEY_CODE,
    ctrlKey: true,
    altKey: true,
    bubbles: true,
    cancelable: true,
  }
  target.dispatchEvent(new KeyboardEvent('keydown', opts))
  // Keyup hơi trễ để giống thao tác thật (một số app chỉ xử lý khi có đủ down+up)
  setTimeout(() => {
    target.dispatchEvent(new KeyboardEvent('keyup', { ...opts }))
  }, 0)
}

// Gọi từ command (phím tắt extension): bấm phím đã đăng ký → gửi Ctrl+Alt+[
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.type !== 'slidesFitToScreen') return
  dispatchFitToScreen(document)
  dispatchFitToScreen(document.body)
  sendResponse({ ok: true })
  return true
})

// Fallback: bắt trực tiếp Ctrl+Alt+/ trong page (có thể không hoạt động do Slides dùng iframe/trusted)
document.addEventListener(
  'keydown',
  (e) => {
    if (!e.ctrlKey || !e.altKey || !isSlashKey(e)) return
    e.preventDefault()
    e.stopPropagation()
    const target = e.target && typeof e.target.dispatchEvent === 'function' ? e.target : document
    dispatchFitToScreen(target)
    if (target !== document) dispatchFitToScreen(document)
  },
  true
)
