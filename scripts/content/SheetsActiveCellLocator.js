// @ts-check
/// <reference path="./types.js" />

/** @implements {ActiveCellLocator} */
class SheetsActiveCellLocator {
  /** @readonly */
  _activeBorderClass = 'active-cell-border'
  /** @readonly */
  _selectionClass = 'selection'
  /** @readonly */
  _sheetContainerId = 'waffle-grid-container'

  getHighlightRectList() {
    const activeSelectionList = Array.from(
      /** @type {HTMLCollectionOf<HTMLElement>} */ (
        document.getElementsByClassName(this._selectionClass)
      )
    ).filter((element) => element.style.display !== 'none')

    if (activeSelectionList.length) {
      return this._getMultipleHighlightRectList(activeSelectionList)
    }

    return this._getSingleHighlightRectList()
  }

  /**
   * Trả về rect ô hiện tại (chọn đơn).
   * Dùng cạnh trong của 4 border (line xanh Sheets) làm toạ độ để line highlight căn đúng khi dày.
   * @returns {HighlightRect | null}
   */
  getActiveCellRect() {
    const sheetRect = this._getSheetContainerRect()
    const activeBorderList = document.getElementsByClassName(
      this._activeBorderClass
    )

    if (!sheetRect || activeBorderList.length !== 4) {
      return null
    }

    const rects = Array.from(activeBorderList).map((el) => el.getBoundingClientRect())
    const horizontal = rects.filter((r) => r.width >= r.height)
    const vertical = rects.filter((r) => r.height > r.width)

    let topR, bottomR, leftR, rightR
    if (horizontal.length >= 2 && vertical.length >= 2) {
      topR = horizontal.reduce((a, r) => (r.top < a.top ? r : a))
      bottomR = horizontal.reduce((a, r) => (r.bottom > a.bottom ? r : a))
      leftR = vertical.reduce((a, r) => (r.left < a.left ? r : a))
      rightR = vertical.reduce((a, r) => (r.right > a.right ? r : a))
    } else {
      topR = rects.reduce((a, r) => (r.top < a.top ? r : a))
      bottomR = rects.reduce((a, r) => (r.bottom > a.bottom ? r : a))
      leftR = rects.reduce((a, r) => (r.left < a.left ? r : a))
      rightR = rects.reduce((a, r) => (r.right > a.right ? r : a))
    }

    // Cạnh trong làm grid line; căn line highlight tại đây để không lệch khi đổi độ dày.
    const leftLine = leftR.right
    const rightLine = rightR.left
    const topLine = topR.bottom
    const bottomLine = bottomR.top

    return {
      x: leftLine - sheetRect.x,
      y: topLine - sheetRect.y,
      width: rightLine - leftLine,
      height: bottomLine - topLine,
    }
  }

  /**
   * @param {Array<HTMLElement>} activeSelectionList
   * @returns {Array<HighlightRect>}
   */
  _getMultipleHighlightRectList(activeSelectionList) {
    const sheetRect = this._getSheetContainerRect()

    if (!sheetRect) {
      return []
    }

    const tolerance = 2
    const edgeTolerance = 80

    /** @type {Array<HighlightRect>} */
    const activeSelectionRectList = activeSelectionList.map((element) => {
      const { x, y, width, height } = element.getBoundingClientRect()
      return {
        x: x - sheetRect.x,
        y: y - sheetRect.y,
        width: width - 1,
        height: height - 0.5,
      }
    })

    // Nhiều rect ghép lại thành full hàng/cột thì không vẽ
    if (this._isFullRowOrColumnSelection(activeSelectionRectList, sheetRect, tolerance, edgeTolerance)) {
      return []
    }

    // Rect full hàng/cột (chiều rộng/cao bằng sheet)
    const rowOrColumnRectList = activeSelectionRectList.filter(
      (rect) =>
        rect.width >= sheetRect.width - tolerance ||
        rect.height >= sheetRect.height - tolerance
    )

    // Loại bỏ rect trùng vị trí full hàng/cột
    return activeSelectionRectList.filter(
      (rect) =>
        !rowOrColumnRectList.some(({ x, y, width, height }) =>
          height < width
            ? rect.y === y && rect.height === height
            : rect.x === x && rect.width === width
        )
    )
  }

  /**
   * Kiểm tra nhiều rect gộp lại có phải full hàng hoặc full cột không.
   * Full row = (1) vùng chọn bắt đầu sát cạnh trái (contentLeft <= edgeTol) và có hàng trải hết, HOẶC
   *           (2) vùng chọn trải đủ chiều rộng container (selection width >= sheetWidth - edgeTol) và có hàng trải hết — cho sheet ít cột có row header.
   * @param {Array<HighlightRect>} rectList
   * @param {DOMRect} sheetRect
   * @param {number} tolerance
   * @param {number} [edgeTolerance]
   * @returns {boolean}
   */
  _isFullRowOrColumnSelection(rectList, sheetRect, tolerance, edgeTolerance = 80) {
    if (rectList.length === 0) return false

    const contentLeft = Math.min(...rectList.map((r) => r.x))
    const contentRight = Math.max(...rectList.map((r) => r.x + r.width))
    const contentTop = Math.min(...rectList.map((r) => r.y))
    const contentBottom = Math.max(...rectList.map((r) => r.y + r.height))
    const selectionWidth = contentRight - contentLeft
    const selectionHeight = contentBottom - contentTop
    const sheetWidth = sheetRect.width
    const sheetHeight = sheetRect.height

    const rowGroups = this._groupRectsByRow(rectList, tolerance)
    for (const group of rowGroups) {
      const left = Math.min(...group.map((r) => r.x))
      const right = Math.max(...group.map((r) => r.x + r.width))
      const groupSpansSelection = left <= contentLeft + tolerance && right >= contentRight - tolerance
      const startsAtLeft = contentLeft <= edgeTolerance
      const spansSheetWidth = selectionWidth >= sheetWidth - edgeTolerance
      const spansFullWidth =
        groupSpansSelection && (startsAtLeft || spansSheetWidth)
      if (spansFullWidth) return true
    }

    const colGroups = this._groupRectsByColumn(rectList, tolerance)
    for (const group of colGroups) {
      const top = Math.min(...group.map((r) => r.y))
      const bottom = Math.max(...group.map((r) => r.y + r.height))
      const groupSpansSelection = top <= contentTop + tolerance && bottom >= contentBottom - tolerance
      const startsAtTop = contentTop <= edgeTolerance
      const spansSheetHeight = selectionHeight >= sheetHeight - edgeTolerance
      const spansFullHeight =
        groupSpansSelection && (startsAtTop || spansSheetHeight)
      if (spansFullHeight) return true
    }

    return false
  }

  /**
   * @param {Array<HighlightRect>} rectList
   * @param {number} tolerance
   * @returns {Array<Array<HighlightRect>>}
   */
  _groupRectsByRow(rectList, tolerance) {
    return this._groupRectsByOverlap(rectList, tolerance, 'y', 'height')
  }

  /**
   * @param {Array<HighlightRect>} rectList
   * @param {number} tolerance
   * @returns {Array<Array<HighlightRect>>}
   */
  _groupRectsByColumn(rectList, tolerance) {
    return this._groupRectsByOverlap(rectList, tolerance, 'x', 'width')
  }

  /**
   * Nhóm các rect chồng nhau theo một trục (bắc cầu).
   * @param {Array<HighlightRect>} rectList
   * @param {number} tolerance
   * @param {'x' | 'y'} pos
   * @param {'width' | 'height'} size
   * @returns {Array<Array<HighlightRect>>}
   */
  _groupRectsByOverlap(rectList, tolerance, pos, size) {
    const used = new Set()
    const groups = []

    for (let i = 0; i < rectList.length; i++) {
      if (used.has(i)) continue
      const group = [rectList[i]]
      used.add(i)
      let added = true
      while (added) {
        added = false
        for (let j = 0; j < rectList.length; j++) {
          if (used.has(j)) continue
          const b = rectList[j]
          const overlaps = group.some((a) => {
            const aEnd = a[pos] + a[size] + tolerance
            const bEnd = b[pos] + b[size] + tolerance
            return a[pos] < bEnd && b[pos] < aEnd
          })
          if (overlaps) {
            group.push(rectList[j])
            used.add(j)
            added = true
          }
        }
      }
      groups.push(group)
    }
    return groups
  }

  /** @returns {Array<HighlightRect>} */
  _getSingleHighlightRectList() {
    const rect = this.getActiveCellRect()
    if (!rect) return []
    const sheetRect = this._getSheetContainerRect()
    if (!sheetRect) return [rect]
    const edgeTolerance = 80
    const minRowWidth = 120
    const minColHeight = 24
    const fullRow =
      rect.x <= edgeTolerance && rect.width >= minRowWidth
    const fullCol =
      rect.y <= edgeTolerance && rect.height >= minColHeight
    if (fullRow || fullCol) return []
    return [rect]
  }

  _getSheetContainerRect() {
    return document
      .getElementById(this._sheetContainerId)
      ?.getBoundingClientRect()
  }

  /**
   * Vùng nội dung sheet (scrollWidth x scrollHeight) để kiểm tra full row/column.
   * Kết hợp với heuristic "chạm cạnh trái + span max" khi sheet ít cột (scrollWidth = clientWidth).
   * @returns {{ left: number, top: number, right: number, bottom: number } | null}
   */
  _getVisibleSheetBounds() {
    const el = document.getElementById(this._sheetContainerId)
    if (!el) return null
    return {
      left: 0,
      top: 0,
      right: el.scrollWidth,
      bottom: el.scrollHeight,
    }
  }

  getSheetContainerStyle() {
    const { x, y, width, height } = this._getSheetContainerRect() || {}

    return {
      left: `${x}px`,
      top: `${y}px`,
      width: `${width}px`,
      height: `${height}px`,
    }
  }

  getSheetKey() {
    const { pathname } = location
    return pathname.match(/d\/([^/]*)/)?.[1] || pathname
  }
}
