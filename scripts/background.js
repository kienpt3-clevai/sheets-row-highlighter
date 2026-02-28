importScripts('scripts/content/types.js')

chrome.commands.onCommand.addListener((command) => {
  chrome.storage.local.get(['row', 'column'], (items) => {
    let row = items.row ?? DEFAULT_ROW
    let column = items.column ?? DEFAULT_COLUMN

    switch (command) {
      case 'toggleRow': {
        row = !row
        break
      }
      case 'toggleColumn': {
        column = !column
        break
      }
      case 'toggleBoth': {
        if (row || column) {
          row = false
          column = false
        } else {
          row = true
          column = true
        }
        break
      }
    }

    chrome.storage.local.set({ row, column })

    // Gửi lệnh xuống tab đang active để content script cập nhật highlight ngay
    chrome.tabs.query({ active: true, lastFocusedWindow: true }, (tabs) => {
      if (tabs[0]?.id) {
        chrome.tabs.sendMessage(tabs[0].id, { type: 'applyCommand', row, column }).catch(() => {})
      }
    })

    chrome.runtime.sendMessage({ type: 'commands', row, column }).catch(() => {})
  })
})
