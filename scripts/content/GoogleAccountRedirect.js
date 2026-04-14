// Auto-redirect Google services to use account /u/0/ instead of /u/N/
(function () {
  function redirectIfNeeded(url) {
    if (!url) return false
    try {
      const fullUrl = new URL(String(url), location.href).href
      const match = fullUrl.match(/\/u\/([1-9]\d*)\//);
      if (match) {
        location.replace(fullUrl.replace(/\/u\/[1-9]\d*\//, '/u/0/'))
        return true
      }
    } catch (_) {}
    return false
  }

  // Kiểm tra ngay khi document_start (bắt server-side redirect)
  if (redirectIfNeeded(location.href)) return

  // Patch History API để bắt SPA navigation (Google dùng replaceState để thêm /u/5/)
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
})()
