const defaultRow = true
const defaultColumn = false

chrome.commands.onCommand.addListener((command) => {
  chrome.storage.local.get(['row', 'column'], (items) => {
    let row = items.row ?? defaultRow
    let column = items.column ?? defaultColumn

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

    chrome.runtime
      .sendMessage({ type: 'commands', row, column })
      .catch(() => {})
  })
})
