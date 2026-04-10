// Auto-redirect Google services to use account /u/0/ instead of /u/5/
(function () {
  const url = location.href
  // Match /u/NUMBER/ where NUMBER is not 0
  const match = url.match(/\/u\/([1-9]\d*)\//);
  if (match) {
    const newUrl = url.replace(/\/u\/[1-9]\d*\//, '/u/0/')
    location.replace(newUrl)
  }
})()
