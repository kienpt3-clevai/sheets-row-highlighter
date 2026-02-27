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

    // アクティブセルのブラウザ座標
    const cellLeft = sheetRect.x + activeRect.x
    const cellRight = cellLeft + activeRect.width
    const cellTop = sheetRect.y + activeRect.y
    const cellBottom = cellTop + activeRect.height

    /**
     * ヘッダーコンテナ内から、アクティブセルに対応する
     * ヘッダーセル（列/行）を幾何的に探す
     * @param {HTMLElement} container
     * @param {'x' | 'y'} axis
     * @returns {DOMRect | null}
     */
    const findHeaderCellInContainer = (container, axis) => {
      /** @type {{rect: DOMRect, score: number, distance: number} | null} */
      let best = null

      const elements = /** @type {NodeListOf<HTMLElement>} */ (
        container.querySelectorAll('div, span')
      )

      elements.forEach((el) => {
        const rect = el.getBoundingClientRect()
        if (!rect || rect.width <= 0 || rect.height <= 0) {
          return
        }

        if (axis === 'x') {
          // 列ヘッダー: アクティブセルと X 方向で重なる要素の中から、
          // シート本体のすぐ上にあるものを優先する
          const overlapX =
            Math.min(cellRight, rect.right) - Math.max(cellLeft, rect.left)
          if (overlapX <= 0) return

          const distance = Math.abs(rect.bottom - sheetRect.y)
          const score = overlapX

          if (
            !best ||
            score > best.score ||
            (score === best.score && distance < best.distance)
          ) {
            best = { rect, score, distance }
          }
        } else {
          // 行ヘッダー: アクティブセルと Y 方向で重なる要素の中から、
          // シート本体のすぐ左にあるものを優先する
          const overlapY =
            Math.min(cellBottom, rect.bottom) - Math.max(cellTop, rect.top)
          if (overlapY <= 0) return

          const distance = Math.abs(rect.right - sheetRect.x)
          const score = overlapY

          if (
            !best ||
            score > best.score ||
            (score === best.score && distance < best.distance)
          ) {
            best = { rect, score, distance }
          }
        }
      })

      return best ? best.rect : null
    }

    const colHeaderCellRect = findHeaderCellInContainer(
      colHeaderContainer,
      'x'
    )
    const rowHeaderCellRect = findHeaderCellInContainer(
      rowHeaderContainer,
      'y'
    )

    /** @type {Array<HighlightRect>} */
    const result = []

    if (colHeaderCellRect) {
      // 列ヘッダー（例: A/B/C/D...）- セル単位の Rect を使用
      result.push({
        x: colHeaderCellRect.x,
        y: colHeaderCellRect.y,
        width: colHeaderCellRect.width,
        height: colHeaderCellRect.height,
      })
    } else {
      // フォールバック: コンテナを使うが、高さは 1 行分程度に制限
      const colHeaderRect = colHeaderContainer.getBoundingClientRect()
      result.push({
        x: cellLeft,
        y: colHeaderRect.y,
        width: activeRect.width,
        height: Math.min(colHeaderRect.height, activeRect.height * 1.2),
      })
    }

    if (rowHeaderCellRect) {
      // 行ヘッダー（例: 3/12...）- セル単位の Rect を使用
      result.push({
        x: rowHeaderCellRect.x,
        y: rowHeaderCellRect.y,
        width: rowHeaderCellRect.width,
        height: rowHeaderCellRect.height,
      })
    } else {
      // フォールバック: 左側のヘッダーバンドだけをカバーする幅に絞る
      const rowHeaderRect = rowHeaderContainer.getBoundingClientRect()
      const fallbackWidth = Math.min(rowHeaderRect.width, 40)
      result.push({
        x: rowHeaderRect.x,
        y: cellTop,
        width: fallbackWidth,
        height: activeRect.height,
      })
    }

    return result
  }

  getSheetKey() {
    const { pathname } = location
    return pathname.match(/d\/([^/]*)/)?.[1] || pathname
  }
}
