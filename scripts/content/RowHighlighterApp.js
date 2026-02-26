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

    /** @type {Array<{isRow: boolean, left: string, top: string, width: string, height: string, isCell?: boolean}>} */
    let highlightTaskList = []

    const hasSingleActiveRect =
      !!activeCellRect && rectList.length === 1

    if (hasSingleActiveRect) {
      const { x, y, width, height } = activeCellRect

      if (this.isRowEnabled) {
        const rowTop = y + alignOffset
        const rowHeight = Math.max(0, height - alignOffset)

        // 左側（アクティブセルの左まで）
        if (x > 0) {
          highlightTaskList.push({
            isRow: true,
            left: '0px',
            top: `${rowTop}px`,
            width: `${x}px`,
            height: `${rowHeight}px`,
          })
        }

        // 右側（アクティブセルの右からシート終わりまで）
        highlightTaskList.push({
          isRow: true,
          left: `${x + width}px`,
          top: `${rowTop}px`,
          width: `calc(100% - ${x + width}px)`,
          height: `${rowHeight}px`,
        })
      }

      if (this.isColEnabled) {
        const colLeft = x + alignOffset
        const colWidth = Math.max(0, width - alignOffset)

        // 上側（アクティブセルの上まで）
        if (y > 0) {
          highlightTaskList.push({
            isRow: false,
            left: `${colLeft}px`,
            top: '0px',
            width: `${colWidth}px`,
            height: `${y}px`,
          })
        }

        // 下側（アクティブセルの下からシート終わりまで）
        highlightTaskList.push({
          isRow: false,
          left: `${colLeft}px`,
          top: `${y + height}px`,
          width: `${colWidth}px`,
          height: `calc(100% - ${y + height}px)`,
        })
      }

      // 交差セル自体も少し強めにハイライト（行＋列が両方有効な場合）
      if (this.isRowEnabled && this.isColEnabled) {
        highlightTaskList.push({
          isRow: true,
          isCell: true,
          left: `${x}px`,
          top: `${y}px`,
          width: `${width}px`,
          height: `${height}px`,
        })
      }
    } else {
      highlightTaskList = (
        this.isRowEnabled
          ? this._mergeRectList(rectList, 'y').map(({ height, y }) => ({
              isRow: true,
              left: '0px',
              top: `${y + alignOffset}px`,
              width: '100%',
              height: `${Math.max(0, height - alignOffset)}px`,
            }))
          : []
      ).concat(
        this.isColEnabled
          ? this._mergeRectList(rectList, 'x').map(({ width, x }) => ({
              isRow: false,
              left: `${x + alignOffset}px`,
              top: '0px',
              width: `${Math.max(0, width - alignOffset)}px`,
              height: '100%',
            }))
          : []
      )
    }

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
      const { isRow, isCell, ...box } = task

      /** @type {Partial<CSSStyleDeclaration>} */
      let border
      /** @type {string} */
      let opacity = this.opacity
      /** @type {string} */
      let backgroundColor = 'transparent'

      if (isCell) {
        border = {
          borderTop: 'none',
          borderBottom: 'none',
          borderLeft: 'none',
          borderRight: 'none',
        }
        backgroundColor = this.backgroundColor
        opacity = '0.2'
      } else {
        border =
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
      }

      Object.assign(element.style, {
        position: 'absolute',
        pointerEvents: 'none',
        display: 'block',
        backgroundColor,
        opacity,
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
