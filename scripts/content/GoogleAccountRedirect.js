// Auto-redirect Google services to use account /u/0/ instead of /u/N/
(function () {
  function redirectIfNeeded(url) {
    if (!url) return false
    try {
      const fullUrl = new URL(String(url), location.href).href
      let newUrl = fullUrl

      // Pattern 1: /u/N/ trong path (docs, sheets, slides)
      newUrl = newUrl.replace(/\/u\/[1-9]\d*\//, '/u/0/')

      // Pattern 2: authuser=N trong query string (meet)
      newUrl = newUrl.replace(/([?&]authuser=)[1-9]\d*(&|$)/, '$10$2')

      if (newUrl !== fullUrl) {
        location.replace(newUrl)
        return true
      }
    } catch (_) {}
    return false
  }

  // Tự động click OK trên dialog xác nhận tài khoản Google
  // Selector chính xác: button[name="ok"] bên trong .active-account-dialog
  function tryAutoClickOk() {
    if (!document.body) return false
    const btn = document.querySelector('.active-account-dialog button[name="ok"]')
    if (btn && btn.offsetParent !== null) {
      btn.click()
      return true
    }
    return false
  }

  function startObserver() {
    const target = document.body || document.documentElement
    if (!target) {
      document.addEventListener('DOMContentLoaded', startObserver, { once: true })
      return
    }
    const observer = new MutationObserver(tryAutoClickOk)
    observer.observe(target, { childList: true, subtree: true })
    tryAutoClickOk() // thử ngay nếu dialog đã render
  }

  // Kiểm tra ngay khi document_start (bắt server-side redirect)
  if (redirectIfNeeded(location.href)) return

  // Patch History API để bắt SPA navigation (Google dùng replaceState để thêm /u/5/ hoặc authuser=5)
  const _pushState = history.pushState.bind(history)
  const _replaceState = history.replaceState.bind(history)

  history.pushState = function (state, title, url) {
    if (url && redirectIfNeeded(String(url))) return
    return _pushState(state, title, url)
  }

  history.replaceState = function (state, title, url) {
    if (url && redirectIfNeeded(String(url))) return
    return _replaceState(state, title, url)
  }

  // Bắt back/forward navigation
  window.addEventListener('popstate', function () {
    redirectIfNeeded(location.href)
  })

  // Bắt dialog xác nhận tài khoản (xuất hiện sau khi load trang)
  startObserver()
})()
