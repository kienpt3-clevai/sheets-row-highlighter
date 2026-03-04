// @ts-check
// Google Slides: Ctrl+Alt+/ → gửi Ctrl+Alt+[ (fit to screen)

const FIT_KEY = '['
const FIT_CODE = 'BracketLeft'
const FIT_KEY_CODE = 219

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
  target.dispatchEvent(new KeyboardEvent('keyup', { ...opts }))
}

document.addEventListener(
  'keydown',
  (e) => {
    if (!e.ctrlKey || !e.altKey || e.key !== '/') return
    e.preventDefault()
    e.stopPropagation()
    dispatchFitToScreen(document)
  },
  true
)
