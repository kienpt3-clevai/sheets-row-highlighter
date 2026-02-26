const defaultRow = true
const defaultColumn = false
const zoomPresets = [50, 75, 90, 100, 125]

const handleCommand = (command) => {
  chrome.storage.local.get(['row', 'column', 'zoomIndex'], (items) => {
    let row = items.row ?? defaultRow
    let column = items.column ?? defaultColumn
    let zoomIndex = typeof items.zoomIndex === 'number' ? items.zoomIndex : 0

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
      case 'cycleZoomOut': {
        zoomIndex =
          (zoomIndex - 1 + zoomPresets.length) % zoomPresets.length
        break
      }
      case 'cycleZoomIn': {
        zoomIndex = (zoomIndex + 1) % zoomPresets.length
        break
      }
    }

    chrome.storage.local.set({ row, column, zoomIndex })

    chrome.runtime
      .sendMessage({
        type: 'commands',
        row,
        column,
      })
      .catch(() => {})

    if (command === 'cycleZoomOut' || command === 'cycleZoomIn') {
      const zoomLevel = zoomPresets[zoomIndex]
      chrome.runtime
        .sendMessage({
          type: 'zoomCommand',
          zoomLevel,
        })
        .catch(() => {})
    }
  })
}

chrome.commands.onCommand.addListener((command) => {
  handleCommand(command)
})

chrome.runtime.onMessage.addListener((message) => {
  if (message?.type === 'cycleZoomOut' || message?.type === 'cycleZoomIn') {
    handleCommand(message.type)
  }
})
