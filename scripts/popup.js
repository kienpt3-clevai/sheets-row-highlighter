const sendMessageToActiveTab = async (message) => {
  const [tab] = await chrome.tabs.query({
    active: true,
    lastFocusedWindow: true,
  })

  return chrome.tabs.sendMessage(tab.id, message)
}

window.addEventListener('load', () => {
  const opacityInput = document.getElementById('opacity')
  const rowInput = document.getElementById('row')
  const columnInput = document.getElementById('column')
  const lineSizeInput = document.getElementById('lineSize')
  const resetButton = document.getElementById('reset')
  const zoomOutBtn = document.getElementById('zoomOutBtn')
  const zoomInBtn = document.getElementById('zoomInBtn')

  const defaultColor = '#c2185b'
  const defaultOpacity = '1'
  const defaultRow = true
  const defaultColumn = true
  const defaultLineSize = 2

  const customColors = [
    '#0e65eb',
    '#1565c0',
    '#1b5e20',
    '#5c0eec',
    '#6a1b9a',
    '#c2185b',
    '#ec0e2f',
    '#ec930e',
  ].sort()

  const hueb = new Huebee('#color', {
    notation: 'hex',
    customColors,
    shades: 0,
    hues: 4,
  })

  // 設定保存
  const save = () => {
    const color = hueb.color
    const opacity = Math.min(
      Math.max(parseFloat(opacityInput.value, 10) || 1, 0.1),
      1
    )
    const row = rowInput.checked
    const column = columnInput.checked
    const lineSize = Math.min(
      Math.max(parseFloat(lineSizeInput.value, 10) || defaultLineSize, 0.5),
      5
    )

    chrome.storage.local.set(
      { color, opacity, row, column, lineSize },
      () => {
        sendMessageToActiveTab({ type: 'settingsUpdated' }).catch(() => {})
      }
    )
  }

  // 設定リセット
  resetButton.addEventListener('click', () => {
    hueb.off('change', save)

    hueb.setColor(defaultColor)
    opacityInput.value = defaultOpacity
    rowInput.checked = defaultRow
    columnInput.checked = defaultColumn
    lineSizeInput.value = defaultLineSize

    chrome.storage.local.set(
      {
        color: defaultColor,
        opacity: defaultOpacity,
        row: defaultRow,
        column: defaultColumn,
        lineSize: defaultLineSize,
      },
      () => {
        sendMessageToActiveTab({ type: 'settingsUpdated' }).catch(() => {})
      }
    )

    hueb.on('change', save)
  })

  const cycleColor = (delta) => {
    const current = hueb.color?.toLowerCase() ?? ''
    const idx = customColors.findIndex((c) => c.toLowerCase() === current)
    const len = customColors.length
    const nextIdx = idx < 0 ? 0 : (idx + delta + len) % len
    hueb.setColor(customColors[nextIdx])
    save()
  }

  document.getElementById('prevColor').addEventListener('click', () => cycleColor(-1))
  document.getElementById('nextColor').addEventListener('click', () => cycleColor(1))

  if (zoomOutBtn) {
    zoomOutBtn.addEventListener('click', () => {
      chrome.runtime.sendMessage({ type: 'cycleZoomOut' }).catch(() => {})
    })
  }

  if (zoomInBtn) {
    zoomInBtn.addEventListener('click', () => {
      chrome.runtime.sendMessage({ type: 'cycleZoomIn' }).catch(() => {})
    })
  }

  // 設定読み込み
  chrome.storage.local.get(
    ['color', 'opacity', 'row', 'column', 'lineSize'],
    (items) => {
      hueb.setColor(items.color ?? defaultColor)
      opacityInput.value = items.opacity ?? defaultOpacity
      rowInput.checked = items.row ?? defaultRow
      columnInput.checked = items.column ?? defaultColumn
      lineSizeInput.value = items.lineSize ?? defaultLineSize

      hueb.on('change', save)
      opacityInput.addEventListener('change', save)
      rowInput.addEventListener('change', save)
      columnInput.addEventListener('change', save)
      lineSizeInput.addEventListener('change', save)
    }
  )

  // ショートカット入力のメッセージを受け取ったとき
  chrome.runtime.onMessage.addListener((request) => {
    if (request.type !== 'commands') return

    rowInput.checked = request.row
    columnInput.checked = request.column
  })
})
