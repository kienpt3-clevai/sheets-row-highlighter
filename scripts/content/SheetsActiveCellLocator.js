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

    /** @type {Array<HighlightRect>} */
    const activeSelectionRectList = activeSelectionList.map((element) => {
      const { x, y, width, height } = element.getBoundingClientRect()
      return {
        x: Math.ceil(x - sheetRect.x),
        y: Math.ceil(y - sheetRect.y),
        width: Math.ceil(width),
        height: Math.ceil(height),
      }
    })

    // Nhiều rect ghép lại thành full hàng/cột thì không vẽ
    if (this._isFullRowOrColumnSelection(activeSelectionRectList, sheetRect, tolerance)) {
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
   * @param {Array<HighlightRect>} rectList
   * @param {DOMRect} sheetRect
   * @param {number} tolerance
   * @returns {boolean}
   */
  _isFullRowOrColumnSelection(rectList, sheetRect, tolerance) {
    if (rectList.length === 0) return false

    const rowGroups = this._groupRectsByRow(rectList, tolerance)
    for (const group of rowGroups) {
      const left = Math.min(...group.map((r) => r.x))
      const right = Math.max(...group.map((r) => r.x + r.width))
      if (right - left >= sheetRect.width - tolerance) return true
    }

    const colGroups = this._groupRectsByColumn(rectList, tolerance)
    for (const group of colGroups) {
      const top = Math.min(...group.map((r) => r.y))
      const bottom = Math.max(...group.map((r) => r.y + r.height))
      if (bottom - top >= sheetRect.height - tolerance) return true
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
    return rect ? [rect] : []
  }

  _getSheetContainerRect() {
    return document
      .getElementById(this._sheetContainerId)
      ?.getBoundingClientRect()
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
