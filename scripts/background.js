importScripts('content/types.js')

// Xử lý yêu cầu captureVisibleTab từ content script
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'requestCapture') {
    const windowId = sender.tab?.windowId ?? null
    chrome.tabs.captureVisibleTab(windowId, { format: 'png' }, (dataUrl) => {
      if (chrome.runtime.lastError) {
        sendResponse({ error: chrome.runtime.lastError.message })
      } else {
        sendResponse({ dataUrl })
      }
    })
    return true // giữ channel cho async response
  }
})

chrome.commands.onCommand.addListener((command) => {
  console.log('[background] command:', command)
  if (command === 'captureSelection') {
    chrome.tabs.query({ active: true, lastFocusedWindow: true }, (tabs) => {
      console.log('[background] captureSelection → tab:', tabs[0]?.id, tabs[0]?.url)
      if (tabs[0]?.id) {
        chrome.tabs.sendMessage(tabs[0].id, { type: 'captureSelection' }).catch((e) => {
          console.error('[background] sendMessage error:', e)
        })
      }
    })
    return
  }

  if (command === 'slidesFitToScreen') {
    chrome.tabs.query({ active: true, lastFocusedWindow: true }, (tabs) => {
      if (tabs[0]?.id) {
        chrome.tabs.sendMessage(tabs[0].id, { type: 'slidesFitToScreen' }).catch(() => {})
      }
    })
    return
  }

  if (command === 'toggleRow' || command === 'toggleColumn' || command === 'toggleBoth') {
    chrome.tabs.query({ active: true, lastFocusedWindow: true }, (tabs) => {
      if (tabs[0]?.id) {
        chrome.tabs
          .sendMessage(tabs[0].id, { type: 'applyCommand', command })
          .catch(() => {})
      }
    })
  }
})
