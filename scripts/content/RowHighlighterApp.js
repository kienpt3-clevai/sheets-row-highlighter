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

    /** @type {HTMLElement} */
    this.headerContainer = document.createElement('div')
    this.headerContainer.id = 'rh-header-container'
    document.body.appendChild(this.headerContainer)

    /** @type {Array<HTMLElement>} */
    this.headerElementPool = []

    this.backgroundColor = '#0e65eb'
    this.opacity = '0.1'
    this.lineSize = 2
    this.isRowEnabled = true
    this.isColEnabled = false
  }

  update() {
    const rectList = this.locator.getHighlightRectList()
    /** @type {HighlightRect | null | undefined} */
    const activeCellRect =
      typeof this.locator.getActiveCellRect === 'function'
        ? this.locator.getActiveCellRect()
        : null

    Object.assign(this.appContainer.style, {
      position: 'absolute',
      pointerEvents: 'none',
      overflow: 'hidden',
      ...this.locator.getSheetContainerStyle(),
    })

    const borderWidth = `${this.lineSize}px`
    const borderStyle = `solid ${this.backgroundColor}`
    const alignOffset = this.lineSize
    const halfOffset = this.lineSize / 2

    /** @type {Array<{isRow: boolean, left: string, top: string, width: string, height: string}>} */
    let highlightTaskList = []

    highlightTaskList = (
      this.isRowEnabled
        ? this._mergeRectList(rectList, 'y').map(({ height, y }) => ({
            isRow: true,
            left: '0px',
            top: `${Math.max(0, y - halfOffset)}px`,
            width: '100%',
            height: `${Math.max(0, height - alignOffset)}px`,
          }))
        : []
    ).concat(
      this.isColEnabled
        ? this._mergeRectList(rectList, 'x').map(({ width, x }) => ({
            isRow: false,
            left: `${Math.max(0, x - halfOffset)}px`,
            top: '0px',
            width: `${Math.max(0, width - alignOffset)}px`,
            height: '100%',
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

    // =======================
    // Header highlight (T / 23)
    // =======================
    /** @type {Array<HighlightRect>} */
    const headerRects =
      typeof this.locator.getHeaderHighlightRectList === 'function'
        ? this.locator.getHeaderHighlightRectList()
        : []

    Object.assign(this.headerContainer.style, {
      position: 'absolute',
      pointerEvents: 'none',
      left: '0px',
      top: '0px',
      width: '100%',
      height: '100%',
      overflow: 'hidden',
    })

    const headerDiff = headerRects.length - this.headerElementPool.length

    if (0 < headerDiff) {
      Array.from({ length: headerDiff }).forEach(() => {
        const element = document.createElement('div')
        this.headerElementPool.push(element)
        this.headerContainer.appendChild(element)
      })
    }

    if (headerDiff < 0) {
      this.headerElementPool.slice(headerDiff).forEach((element) => {
        element.style.display = 'none'
      })
    }

    headerRects.forEach((rect, index) => {
      const element = this.headerElementPool[index]
      let { x, y, width, height } = rect

      // headerRects[0]: 列ヘッダー（A/B/C...）
      // headerRects[1]: 行ヘッダー（1/2/3...）
      if (index === 0 && activeCellRect) {
        // Chỉ cao bằng 1 dòng, không kéo xuống toàn bộ vùng freeze
        height = Math.min(height, activeCellRect.height)
      } else if (index === 1) {
        // Chỉ rộng bằng dải số dòng bên trái (~46px),
        // không tràn vào các cột freeze (A, B, ...)
        width = Math.min(width, 46)
      }

      Object.assign(element.style, {
        position: 'absolute',
        pointerEvents: 'none',
        display: 'block',
        backgroundColor: this.backgroundColor,
        opacity: '0.1',
        left: `${x}px`,
        top: `${y}px`,
        width: `${width}px`,
        height: `${height}px`,
        border: 'none',
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
