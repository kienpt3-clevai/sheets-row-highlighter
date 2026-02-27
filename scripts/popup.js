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
  const headerColScaleInput = document.getElementById('headerColScale')
  const headerRowScaleInput = document.getElementById('headerRowScale')
  const resetButton = document.getElementById('reset')

  const defaultColor = '#c2185b'
  const defaultOpacity = '0.8'
  const defaultRow = true
  const defaultColumn = true
  const defaultLineSize = 1.75
  const defaultHeaderColScale = 0.9
  const defaultHeaderRowScale = 1.15

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
    const headerColScale = Math.min(
      Math.max(parseFloat(headerColScaleInput.value, 10) || defaultHeaderColScale, 0.5),
      4
    )
    const headerRowScale = Math.min(
      Math.max(parseFloat(headerRowScaleInput.value, 10) || defaultHeaderRowScale, 0.5),
      4
    )

    chrome.storage.local.set(
      { color, opacity, row, column, lineSize, headerColScale, headerRowScale },
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
    headerColScaleInput.value = defaultHeaderColScale
    headerRowScaleInput.value = defaultHeaderRowScale

    chrome.storage.local.set(
      {
        color: defaultColor,
        opacity: defaultOpacity,
        row: defaultRow,
        column: defaultColumn,
        lineSize: defaultLineSize,
        headerColScale: defaultHeaderColScale,
        headerRowScale: defaultHeaderRowScale,
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

  // 設定読み込み
  chrome.storage.local.get(
    ['color', 'opacity', 'row', 'column', 'lineSize', 'headerColScale', 'headerRowScale'],
    (items) => {
      hueb.setColor(items.color ?? defaultColor)
      opacityInput.value = items.opacity ?? defaultOpacity
      rowInput.checked = items.row ?? defaultRow
      columnInput.checked = items.column ?? defaultColumn
      lineSizeInput.value = items.lineSize ?? defaultLineSize
      headerColScaleInput.value = items.headerColScale ?? defaultHeaderColScale
      headerRowScaleInput.value = items.headerRowScale ?? defaultHeaderRowScale

      hueb.on('change', save)
      opacityInput.addEventListener('change', save)
      rowInput.addEventListener('change', save)
      columnInput.addEventListener('change', save)
      lineSizeInput.addEventListener('change', save)
      headerColScaleInput.addEventListener('change', save)
      headerRowScaleInput.addEventListener('change', save)
    }
  )

  // ショートカット入力のメッセージを受け取ったとき
  chrome.runtime.onMessage.addListener((request) => {
    if (request.type !== 'commands') return

    rowInput.checked = request.row
    columnInput.checked = request.column
  })
})
