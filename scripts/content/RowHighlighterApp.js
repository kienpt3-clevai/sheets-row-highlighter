// @ts-check
/// <reference path="./types.js" />

class RowHighlighterApp {
  /**
   * @param {HTMLElement} appContainer
   * @param {ActiveCellLocator} locator
   */
  constructor(appContainer, locator) {
    this.appContainer = appContainer
    this.locator = locator

    /** @type {Array<HTMLElement>} */
    this.elementPool = []

    this.backgroundColor = '#0e65eb'
    this.rowLineColor = this.backgroundColor
    this.colLineColor = this.backgroundColor
    this.rowFillColor = this.backgroundColor
    this.colFillColor = this.backgroundColor
    this.opacity = '1'
    this.rowLineSize = 3.25
    this.colLineSize = 3.25
    // Opacity fill riêng cho row/col (popup: Row Opacity / Col Opacity)
    this.rowFillOpacity = 0.05
    this.colFillOpacity = 0.05
    this.isRowEnabled = true
    this.isColEnabled = false
    this.fillRowEnabled = true
    this.fillColEnabled = true
    /** Inset (px): offset from grid line; 0 = tâm line trùng line xanh dương (cần cho line dày 4.25) */
    this.lineInsetLeft = 0
    this.lineInsetRight = 0
    this.lineInsetTop = 0
    this.lineInsetBottom = 0
    /** offsetHenry (px): trái -0.75, trên -0.5, dưới +0.75, phải +1 */
    this.offsetHenryLeft = -0.75
    this.offsetHenryTop = -0.5
    this.offsetHenryBottom = 0.75
    this.offsetHenryRight = 1
  }

  update() {
    const rectList = this.locator.getHighlightRectList()

    Object.assign(this.appContainer.style, {
      position: 'absolute',
      pointerEvents: 'none',
      overflow: 'visible',
      ...this.locator.getSheetContainerStyle(),
    })

    // Tâm line trùng line xanh dương; dày đều về 2 phía (cả 4 line: trên, dưới, trái, phải)
    const halfRowLine = this.rowLineSize / 2
    const halfColLine = this.colLineSize / 2

    /** @type {Array<{isRow?: boolean, isCellFill?: boolean, left: string, top: string, width: string, height: string}>} */
    let highlightTaskList = []

    const insL = this.lineInsetLeft
    const insR = this.lineInsetRight
    const insT = this.lineInsetTop
    const insB = this.lineInsetBottom
    const oHL = this.offsetHenryLeft
    const oHT = this.offsetHenryTop
    const oHB = this.offsetHenryBottom
    const oHR = this.offsetHenryRight

    const rowBands = this.isRowEnabled
      ? this._mergeRectList(rectList, 'y').map(({ height, y }) => ({
          left: `${insL + oHL}px`,
          top: `${y - halfRowLine + oHT}px`,
          width: `calc(100% - ${insL + insR - oHL - oHR}px)`,
          height: `${height - insB + this.rowLineSize + oHB - oHT}px`,
        }))
      : []
    const colBands = this.isColEnabled
      ? this._mergeRectList(rectList, 'x').map(({ width, x }) => ({
          left: `${x + insL - halfColLine + oHL}px`,
          top: `${insT + oHT}px`,
          width: `${width - insL - insR + this.colLineSize + oHR - oHL}px`,
          height: `calc(100% - ${insT + insB - oHB + oHT}px)`,
        }))
      : []

    if (rectList.length > 0) {
      if (this.fillRowEnabled && this.rowFillOpacity > 0) {
        for (const box of rowBands) {
          highlightTaskList.push({
            isCellFill: true,
            isRow: true,
            fillOpacity: this.rowFillOpacity,
            ...box,
          })
        }
      }
      if (this.fillColEnabled && this.colFillOpacity > 0) {
        for (const box of colBands) {
          highlightTaskList.push({
            isCellFill: true,
            isRow: false,
            fillOpacity: this.colFillOpacity,
            ...box,
          })
        }
      }
    }

    const lineTasks = [
      ...rowBands.map((box) => ({ ...box, isRow: true })),
      ...colBands.map((box) => ({ ...box, isRow: false })),
    ]
    highlightTaskList = highlightTaskList.concat(lineTasks)

    const diff = highlightTaskList.length - this.elementPool.length

    if (0 < diff) {
      Array.from({ length: diff }).forEach(() => {
        const element = document.createElement('div')
        this.elementPool.push(element)
        this.appContainer.appendChild(element)
      })
    }

    if (diff < 0) {
      this.elementPool.slice(diff).forEach((element) => {
        element.style.display = 'none'
      })
    }

    highlightTaskList.forEach((task, index) => {
      const element = this.elementPool[index]
      const { isRow, isCellFill, fillOpacity, ...box } = task

      if (isCellFill) {
        const fillColor = isRow === true ? this.rowFillColor : this.colFillColor
        Object.assign(element.style, {
          position: 'absolute',
          pointerEvents: 'none',
          display: 'block',
          boxSizing: 'border-box',
          backgroundColor: fillColor,
          opacity: String(
            typeof fillOpacity === 'number' ? fillOpacity : 0
          ),
          border: 'none',
          ...box,
        })
        return
      }

      const lineColor = isRow === true ? this.rowLineColor : this.colLineColor
      const borderStyle = `solid ${lineColor}`
      const border =
        isRow === true
          ? {
              borderTop: `${this.rowLineSize}px ${borderStyle}`,
              borderBottom: `${this.rowLineSize}px ${borderStyle}`,
              borderLeft: 'none',
              borderRight: 'none',
            }
          : {
              borderLeft: `${this.colLineSize}px ${borderStyle}`,
              borderRight: `${this.colLineSize}px ${borderStyle}`,
              borderTop: 'none',
              borderBottom: 'none',
            }

      Object.assign(element.style, {
        position: 'absolute',
        pointerEvents: 'none',
        display: 'block',
        boxSizing: 'border-box',
        backgroundColor: 'transparent',
        opacity: this.opacity,
        ...box,
        ...border,
      })
    })
  }

  /**
   * Gộp các rect chồng nhau theo trục dim, trả về danh sách mới.
   * @param {Array<HighlightRect>} rectList
   * @param {'x' | 'y'} dim
   * @returns {Array<HighlightRect>}
   */
  _mergeRectList(rectList, dim) {
    return [...rectList]
      .sort((a, b) => a[dim] - b[dim])
      .reduce((acc, rect) => {
        const prevRect = acc[acc.length - 1]
        const dimSize = dim === 'x' ? 'width' : 'height'

        if (!prevRect || prevRect[dim] + prevRect[dimSize] < rect[dim]) {
          acc.push({ ...rect })
          return acc
        }

        prevRect[dimSize] = Math.max(
          prevRect[dimSize],
          rect[dim] + rect[dimSize] - prevRect[dim]
        )

        return acc
      }, /** @type {Array<HighlightRect>} */ ([]))
  }
}
