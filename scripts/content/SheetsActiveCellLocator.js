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
   * Dùng cạnh trong của 4 border (cạnh chạm ô = line xanh dương) làm toạ độ, để line highlight căn tâm đúng cả khi dày (4.25).
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

    // Dùng tâm từng border làm vị trí grid line (line highlight căn tâm sẽ trùng dù line dày hay mỏng)
    const leftLine = leftR.x + leftR.width / 2
    const rightLine = rightR.x + rightR.width / 2
    const topLine = topR.y + topR.height / 2
    const bottomLine = bottomR.y + bottomR.height / 2

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
    const rowHeaderCellRect = findHeaderCellInContainer(rowHeaderContainer, 'y')

    /** @type {Array<HighlightRect>} */
    const result = []

    // 列ヘッダー（例: A/B/C/D...）
    // freeze 有無や行の高さに関わらず、視覚的な「1 dòng header」分だけを塗る
    const colHeaderRect = colHeaderContainer.getBoundingClientRect()

    // Thử đọc zoom từ toolbar; nếu không được, fallback theo tỉ lệ chiều cao container
    const zoomFromToolbar = typeof this._getZoomScale === 'function' ? this._getZoomScale() : 0

    // Ghi lại chiều cao container ở lần đầu để làm mốc 100%
    if (!this._baseHeaderContainerHeight) {
      this._baseHeaderContainerHeight = colHeaderRect.height
    }
    const heightScale =
      this._baseHeaderContainerHeight > 0
        ? colHeaderRect.height / this._baseHeaderContainerHeight
        : 1

    const zoomScale =
      zoomFromToolbar && Number.isFinite(zoomFromToolbar) && zoomFromToolbar > 0
        ? zoomFromToolbar
        : heightScale

    const baseHeaderHeight = 26 * zoomScale // px, scale theo zoom thực tế
    const colBandInset = 0 // dịch sát về bên trái hơn một chút
    const colBandWidth = Math.max(0, activeRect.width - colBandInset * 2)
    result.push({
      x: cellLeft + colBandInset,
      // Bắt đầu ngay từ mép trên vùng dữ liệu (chữ cột AH),
      // không ăn vào hàng dấu "+" phía trên.
      y: sheetRect.y,
      width: colBandWidth,
      height: Math.min(colHeaderRect.height, baseHeaderHeight),
    })

    // 行ヘッダー（例: 2/3/12...）
    const rowHeaderRect = rowHeaderContainer.getBoundingClientRect()
    const bandWidth = Math.min(rowHeaderRect.width, 40)
    const rowX = rowHeaderRect.x
    const rowY = rowHeaderCellRect ? rowHeaderCellRect.y : cellTop
    const rowHeight = rowHeaderCellRect
      ? rowHeaderCellRect.height
      : activeRect.height

    result.push({
      x: rowX,
      y: rowY,
      width: bandWidth,
      height: rowHeight,
    })

    return result
  }

  getSheetKey() {
    const { pathname } = location
    return pathname.match(/d\/([^/]*)/)?.[1] || pathname
  }

  /**
   * Google Sheets のツールバーからズーム倍率を推定する
   * 例: "75%" / "100%" / "150%"
   * @returns {number} ズーム倍率 (1.0 = 100%)
   */
  _getZoomScale() {
    try {
      /** @type {HTMLElement | null} */
      const zoomButton =
        document.querySelector('[aria-label="Zoom"]') ||
        document.querySelector('[aria-label="Thu phóng"]') ||
        document.querySelector('[aria-label="Phóng to thu nhỏ"]')

      if (!zoomButton) return 1

      const text =
        zoomButton.textContent ||
        /** @type {HTMLElement | null} */ (
          zoomButton.querySelector('.goog-flat-menu-button-caption')
        )?.textContent ||
        ''

      const match = text.match(/(\d+)\s*%/)
      if (!match) return 1

      const value = Number(match[1])
      if (!Number.isFinite(value) || value <= 0) return 1

      // Giới hạn trong khoảng hợp lý 25% - 400%
      const clamped = Math.max(25, Math.min(400, value))
      return clamped / 100
    } catch {
      return 1
    }
  }
}
