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
    this.opacity = '0.1'
    this.lineSize = 3.25
    this.isRowEnabled = true
    this.isColEnabled = false
    /** Inset (px): line position offset from cell boundaries */
    this.lineInsetLeft = 1
    this.lineInsetRight = 0.5
    this.lineInsetTop = 0
    this.lineInsetBottom = 3.5
  }

  update() {
    const rectList = this.locator.getHighlightRectList()

    Object.assign(this.appContainer.style, {
      position: 'absolute',
      pointerEvents: 'none',
      overflow: 'hidden',
      ...this.locator.getSheetContainerStyle(),
    })

    const borderWidth = `${this.lineSize}px`
    const borderStyle = `solid ${this.backgroundColor}`

    /** @type {Array<{isRow: boolean, left: string, top: string, width: string, height: string}>} */
    let highlightTaskList = []

    const insL = this.lineInsetLeft
    const insR = this.lineInsetRight
    const insT = this.lineInsetTop
    const insB = this.lineInsetBottom

    highlightTaskList = (
      this.isRowEnabled
        ? this._mergeRectList(rectList, 'y').map(({ height, y }) => ({
            isRow: true,
            left: `${insL}px`,
            top: `${Math.max(0, y)}px`,
            width: `calc(100% - ${insL + insR}px)`,
            height: `${Math.max(0, height - insB)}px`,
          }))
        : []
    ).concat(
      this.isColEnabled
        ? this._mergeRectList(rectList, 'x').map(({ width, x }) => ({
            isRow: false,
            left: `${Math.max(0, x + insL)}px`,
            top: `${insT}px`,
            width: `${Math.max(0, width - insL - insR)}px`,
            height: `calc(100% - ${insT + insB}px)`,
          }))
        : []
    )

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
      const { isRow, ...box } = task

      const border =
        isRow === true
          ? {
              borderTop: `${borderWidth} ${borderStyle}`,
              borderBottom: `${borderWidth} ${borderStyle}`,
              borderLeft: 'none',
              borderRight: 'none',
            }
          : {
              borderLeft: `${borderWidth} ${borderStyle}`,
              borderRight: `${borderWidth} ${borderStyle}`,
              borderTop: 'none',
              borderBottom: 'none',
            }

      Object.assign(element.style, {
        position: 'absolute',
        pointerEvents: 'none',
        display: 'block',
        backgroundColor: 'transparent',
        opacity: this.opacity,
        ...box,
        ...border,
      })
    })
  }

  /**
   * 範囲が重なるRectをマージして新しいリストを返す
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
