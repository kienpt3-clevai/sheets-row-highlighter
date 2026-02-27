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
   * 現在のアクティブセルの Rect を返す（単一選択ベース）
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

    const topBorderRect = activeBorderList[0].getBoundingClientRect()
    const leftBorderRect = activeBorderList[3].getBoundingClientRect()

    return {
      x: topBorderRect.x - sheetRect.x,
      y: topBorderRect.y - sheetRect.y,
      width: topBorderRect.width,
      height: leftBorderRect.height,
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

    // 複数rectで full row/column になっているか（Shift+Space / Ctrl+Space）
    if (this._isFullRowOrColumnSelection(activeSelectionRectList, sheetRect, tolerance)) {
      return []
    }

    // 1 rect で full row/column のもの（width/height が sheet 以上）
    const rowOrColumnRectList = activeSelectionRectList.filter(
      (rect) =>
        rect.width >= sheetRect.width - tolerance ||
        rect.height >= sheetRect.height - tolerance
    )

    // 行か列選択と同じ位置のセルを除外して返す
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
   * 複数の rect をまとめたときに full row または full column になっているか
   * @param {Array<HighlightRect>} rectList
   * @param {DOMRect} sheetRect
   * @param {number} tolerance
   * @returns {boolean}
   */
  _isFullRowOrColumnSelection(rectList, sheetRect, tolerance) {
    if (rectList.length === 0) return false

    // 同じ行の rect をマージして幅を計算 → 幅が sheet 以上なら full row
    const rowGroups = this._groupRectsByRow(rectList, tolerance)
    for (const group of rowGroups) {
      const left = Math.min(...group.map((r) => r.x))
      const right = Math.max(...group.map((r) => r.x + r.width))
      if (right - left >= sheetRect.width - tolerance) return true
    }

    // 同じ列の rect をマージして高さを計算 → 高さが sheet 以上なら full column
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
   * Group rects that overlap along one axis (transitively).
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

  /**
   * アクティブセルに対応する列ヘッダーと行ヘッダーの Rect を返す
   * （ブラウザ座標系での位置）
   * @returns {Array<HighlightRect>}
   */
  getHeaderHighlightRectList() {
    const activeRect = this.getActiveCellRect?.()
    const sheetRect = this._getSheetContainerRect()

    if (!activeRect || !sheetRect) {
      return []
    }

    const colHeaderContainer = document.querySelector(
      '.fixed4-inner-container'
    )
    const rowHeaderContainer = document.querySelector(
      '.grid4-inner-container'
    )

    if (!(colHeaderContainer instanceof HTMLElement) || !(rowHeaderContainer instanceof HTMLElement)) {
      return []
    }

    const colHeaderRect = colHeaderContainer.getBoundingClientRect()
    const rowHeaderRect = rowHeaderContainer.getBoundingClientRect()

    // アクティブセルのブラウザ座標
    const cellLeft = sheetRect.x + activeRect.x
    const cellTop = sheetRect.y + activeRect.y

    // ヘッダーの高さは 1 行ぶんだけに制限する
    const headerHeight = Math.max(
      0,
      Math.min(colHeaderRect.height, activeRect.height)
    )

    // 行ヘッダーの幅は「左のヘッダー帯」部分だけに制限する
    let headerWidth = Math.max(0, sheetRect.x - rowHeaderRect.x)
    if (!headerWidth || headerWidth > rowHeaderRect.width) {
      headerWidth = Math.min(rowHeaderRect.width, activeRect.width)
    }

    /** @type {Array<HighlightRect>} */
    const result = []

    // 列ヘッダー（例: F / AR / AU）
    result.push({
      x: cellLeft,
      y: colHeaderRect.y,
      width: activeRect.width,
      height: headerHeight,
    })

    // 行ヘッダー（例: 3）
    result.push({
      x: rowHeaderRect.x,
      y: cellTop,
      width: headerWidth,
      height: activeRect.height,
    })

    return result
  }

  getSheetKey() {
    const { pathname } = location
    return pathname.match(/d\/([^/]*)/)?.[1] || pathname
  }
}
