// @ts-check

/**
 * Chụp vùng chọn Google Sheets thành ảnh PNG, copy vào clipboard.
 *
 * Ưu tiên đọc trực tiếp grid canvas của Sheets (thường render ở 2x DPR nội bộ
 * để crisp khi scroll/zoom) → ảnh sharp gấp đôi mà không cần zoom browser.
 * Fallback: captureVisibleTab nếu không tìm được canvas hoặc canvas bị tainted.
 *
 * Nén: các dòng/cột không có selection → gap 6px + đường phân cách.
 */
class SheetsCapturer {
  constructor() {
    this._GAP_PX = 0
  }

  /**
   * Lấy selection rects ở LOCAL viewport coords (cùng frame với canvas).
   * @returns {Array<{left: number, top: number, width: number, height: number}>}
   */
  _getSelectionRectsLocal() {
    const selections = Array.from(
      document.getElementsByClassName('selection')
    ).filter((el) => el instanceof HTMLElement && el.style.display !== 'none')

    if (selections.length > 0) {
      return selections.map((el) => {
        const r = el.getBoundingClientRect()
        return { left: r.left, top: r.top, width: r.width, height: r.height }
      })
    }

    const borders = Array.from(
      document.getElementsByClassName('active-cell-border')
    )
    if (borders.length === 4) {
      const rects = borders.map((el) => el.getBoundingClientRect())
      const left = Math.min(...rects.map((r) => r.left))
      const top = Math.min(...rects.map((r) => r.top))
      const right = Math.max(...rects.map((r) => r.right))
      const bottom = Math.max(...rects.map((r) => r.bottom))
      return [{ left, top, width: right - left, height: bottom - top }]
    }
    return []
  }

  /**
   * Offset từ frame hiện tại đến top frame (cho captureVisibleTab fallback).
   * @returns {{x: number, y: number}}
   */
  _getFrameOffset() {
    let x = 0
    let y = 0
    let win = window
    try {
      while (win !== window.top) {
        const frame = win.frameElement
        if (!frame) break
        const r = frame.getBoundingClientRect()
        x += r.left
        y += r.top
        win = /** @type {Window} */ (win.parent)
      }
    } catch (_) {}
    return { x, y }
  }

  /**
   * Tìm tất cả grid canvases (tìm toàn document vì Sheets có nhiều pane:
   * frozen rows, frozen cols, scrollable, corner).
   * @returns {Array<{canvas: HTMLCanvasElement, rect: DOMRect, scaleX: number, scaleY: number}>}
   */
  _findCanvases() {
    const canvases = Array.from(document.querySelectorAll('canvas'))
    const result = []
    for (const canvas of canvases) {
      const rect = canvas.getBoundingClientRect()
      if (rect.width < 10 || rect.height < 10) continue
      if (canvas.width < 1 || canvas.height < 1) continue
      const scaleX = canvas.width / rect.width
      const scaleY = canvas.height / rect.height
      result.push({ canvas, rect, scaleX, scaleY })
    }
    return result
  }

  /**
   * Tìm canvas chứa rect (viewport coords cùng frame).
   */
  _findCanvasFor(canvases, rect) {
    const tol = 2
    return canvases.find(
      (c) =>
        rect.left >= c.rect.left - tol &&
        rect.top >= c.rect.top - tol &&
        rect.left + rect.width <= c.rect.right + tol &&
        rect.top + rect.height <= c.rect.bottom + tol
    )
  }

  /**
   * Ẩn overlay cho screenshot fallback.
   * @returns {() => void}
   */
  _hideOverlays() {
    /** @type {Array<{el: HTMLElement, val: string}>} */
    const saved = []
    const hide = (/** @type {Element} */ el) => {
      if (el instanceof HTMLElement) {
        saved.push({ el, val: el.style.visibility })
        el.style.visibility = 'hidden'
      }
    }
    const rh = document.getElementById('rh-app-container')
    if (rh) hide(rh)
    document.querySelectorAll('.selection').forEach(hide)
    document.querySelectorAll('.active-cell-border').forEach(hide)
    return () => saved.forEach(({ el, val }) => (el.style.visibility = val))
  }

  /** @returns {Promise<void>} */
  _waitForRepaint() {
    return new Promise((resolve) =>
      requestAnimationFrame(() => setTimeout(resolve, 50))
    )
  }

  /** @returns {Promise<string>} */
  _requestCapture() {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({ type: 'requestCapture' }, (response) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message))
        } else if (response?.dataUrl) {
          resolve(response.dataUrl)
        } else {
          reject(new Error(response?.error || 'No screenshot data'))
        }
      })
    })
  }

  /**
   * @param {string} src
   * @returns {Promise<HTMLImageElement>}
   */
  _loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image()
      img.onload = () => resolve(img)
      img.onerror = reject
      img.src = src
    })
  }

  /**
   * Gộp khoảng interval chồng nhau.
   * @param {Array<[number, number]>} intervals
   * @returns {Array<[number, number]>}
   */
  _mergeIntervals(intervals) {
    if (!intervals.length) return []
    const sorted = [...intervals].sort((a, b) => a[0] - b[0])
    const merged = [[sorted[0][0], sorted[0][1]]]
    for (let i = 1; i < sorted.length; i++) {
      const last = merged[merged.length - 1]
      if (sorted[i][0] <= last[1] + 2) {
        last[1] = Math.max(last[1], sorted[i][1])
      } else {
        merged.push([sorted[i][0], sorted[i][1]])
      }
    }
    return /** @type {Array<[number, number]>} */ (merged)
  }

  /**
   * Build map interval → offset trong output đã nén.
   */
  _buildCompressedMap(intervals) {
    let offset = 0
    const map = []
    for (const [start, end] of intervals) {
      map.push({ start, end, offset, size: end - start })
      offset += end - start + this._GAP_PX
    }
    return {
      map,
      totalSize: offset - (intervals.length > 0 ? this._GAP_PX : 0),
    }
  }

  _findInterval(map, pos, size) {
    return map.find((i) => pos >= i.start - 2 && pos + size <= i.end + 2)
  }

  /**
   * (Đã bỏ) Trước đây vẽ đường nét đứt giữa các nhóm đã nén.
   * Giờ chỉ giữ gap trắng, không kẻ line.
   */
  _drawSeparators(_ctx, _yMap, _xMap, _outScale, _canvasW, _canvasH) {}

  /**
   * Render qua cách đọc canvas nội bộ của Sheets.
   * @returns {HTMLCanvasElement | null} Null nếu không làm được (fallback).
   */
  _renderViaCanvases(rects, canvases) {
    if (canvases.length === 0) return null

    // Map từng rect → canvas chứa nó
    const mapping = rects.map((r) => ({
      rect: r,
      source: this._findCanvasFor(canvases, r),
    }))
    const missing = mapping.filter((m) => !m.source)
    if (missing.length > 0) {
      console.warn(
        '[SheetsCapturer] rects without canvas:',
        missing.map((m) => m.rect),
        'available canvases:',
        canvases.map((c) => ({
          left: Math.round(c.rect.left),
          top: Math.round(c.rect.top),
          right: Math.round(c.rect.right),
          bottom: Math.round(c.rect.bottom),
          scale: c.scaleX.toFixed(2),
        }))
      )
      return null
    }

    // Output scale = max scale của canvas (để tận dụng hết pixel)
    const outScale = Math.max(...canvases.map((c) => c.scaleX))

    const yIntervals = this._mergeIntervals(
      rects.map((r) => [r.top, r.top + r.height])
    )
    const xIntervals = this._mergeIntervals(
      rects.map((r) => [r.left, r.left + r.width])
    )
    const yMap = this._buildCompressedMap(yIntervals)
    const xMap = this._buildCompressedMap(xIntervals)

    const outCanvas = document.createElement('canvas')
    outCanvas.width = Math.max(1, Math.round(xMap.totalSize * outScale))
    outCanvas.height = Math.max(1, Math.round(yMap.totalSize * outScale))
    const ctx = outCanvas.getContext('2d')
    if (!ctx) return null

    ctx.imageSmoothingEnabled = false
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, outCanvas.width, outCanvas.height)

    for (const { rect, source } of mapping) {
      if (!source) continue
      const yInt = this._findInterval(yMap.map, rect.top, rect.height)
      const xInt = this._findInterval(xMap.map, rect.left, rect.width)
      if (!yInt || !xInt) continue

      const sx = Math.round((rect.left - source.rect.left) * source.scaleX)
      const sy = Math.round((rect.top - source.rect.top) * source.scaleY)
      const sw = Math.round(rect.width * source.scaleX)
      const sh = Math.round(rect.height * source.scaleY)
      const dx = Math.round(
        (xInt.offset + Math.max(0, rect.left - xInt.start)) * outScale
      )
      const dy = Math.round(
        (yInt.offset + Math.max(0, rect.top - yInt.start)) * outScale
      )
      const dw = Math.round(rect.width * outScale)
      const dh = Math.round(rect.height * outScale)

      try {
        ctx.drawImage(source.canvas, sx, sy, sw, sh, dx, dy, dw, dh)
      } catch (e) {
        console.error('[SheetsCapturer] drawImage from canvas failed:', e)
        return null
      }
    }

    this._drawSeparators(ctx, yMap, xMap, outScale, outCanvas.width, outCanvas.height)
    return outCanvas
  }

  /**
   * Fallback: captureVisibleTab + crop (DPR của display).
   * @returns {Promise<HTMLCanvasElement | null>}
   */
  async _renderViaScreenshot(localRects) {
    const offset = this._getFrameOffset()
    const globalRects = localRects.map((r) => ({
      left: r.left + offset.x,
      top: r.top + offset.y,
      width: r.width,
      height: r.height,
    }))

    const restore = this._hideOverlays()
    try {
      await this._waitForRepaint()
      const dataUrl = await this._requestCapture()
      restore()

      const img = await this._loadImage(dataUrl)
      const dpr = window.devicePixelRatio || 1

      const yIntervals = this._mergeIntervals(
        localRects.map((r) => [r.top, r.top + r.height])
      )
      const xIntervals = this._mergeIntervals(
        localRects.map((r) => [r.left, r.left + r.width])
      )
      const yMap = this._buildCompressedMap(yIntervals)
      const xMap = this._buildCompressedMap(xIntervals)

      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1, Math.round(xMap.totalSize * dpr))
      canvas.height = Math.max(1, Math.round(yMap.totalSize * dpr))
      const ctx = canvas.getContext('2d')
      if (!ctx) return null

      ctx.imageSmoothingEnabled = false
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, canvas.width, canvas.height)

      for (let i = 0; i < localRects.length; i++) {
        const r = localRects[i]
        const g = globalRects[i]
        const yInt = this._findInterval(yMap.map, r.top, r.height)
        const xInt = this._findInterval(xMap.map, r.left, r.width)
        if (!yInt || !xInt) continue

        const sx = Math.round(g.left * dpr)
        const sy = Math.round(g.top * dpr)
        const sw = Math.round(g.width * dpr)
        const sh = Math.round(g.height * dpr)
        const dx = Math.round(
          (xInt.offset + Math.max(0, r.left - xInt.start)) * dpr
        )
        const dy = Math.round(
          (yInt.offset + Math.max(0, r.top - yInt.start)) * dpr
        )
        ctx.drawImage(img, sx, sy, sw, sh, dx, dy, sw, sh)
      }

      this._drawSeparators(ctx, yMap, xMap, dpr, canvas.width, canvas.height)
      return canvas
    } catch (e) {
      restore()
      throw e
    }
  }

  /**
   * Copy canvas → clipboard dưới dạng PNG.
   */
  async _copyToClipboard(canvas) {
    const blob = await new Promise((resolve) =>
      canvas.toBlob(resolve, 'image/png')
    )
    if (!blob) throw new Error('toBlob failed')
    await navigator.clipboard.write([
      new ClipboardItem({ 'image/png': blob }),
    ])
  }

  /**
   * @param {string} text
   * @param {boolean} [isError]
   */
  _showToast(text, isError) {
    const toast = document.createElement('div')
    Object.assign(toast.style, {
      position: 'fixed',
      bottom: '24px',
      left: '50%',
      transform: 'translateX(-50%)',
      background: isError ? '#d32f2f' : '#323232',
      color: '#fff',
      padding: '8px 20px',
      borderRadius: '6px',
      fontSize: '13px',
      fontFamily: 'Roboto, Arial, sans-serif',
      zIndex: '99999',
      pointerEvents: 'none',
      transition: 'opacity 0.3s',
      boxShadow: '0 2px 8px rgba(0,0,0,.3)',
    })
    toast.textContent = text
    document.body.appendChild(toast)
    setTimeout(() => {
      toast.style.opacity = '0'
      setTimeout(() => toast.remove(), 300)
    }, 1500)
  }

  /**
   * Entry point.
   */
  async capture() {
    console.log('[SheetsCapturer] capture() called')
    const rects = this._getSelectionRectsLocal()
    console.log('[SheetsCapturer] selection rects:', rects)

    if (rects.length === 0) {
      this._showToast('No selection to capture', true)
      return
    }

    try {
      const canvases = this._findCanvases()
      console.log(
        '[SheetsCapturer] canvases: found=' + canvases.length,
        canvases
          .map(
            (c, i) =>
              `#${i}: canvas=${c.canvas.width}x${c.canvas.height} css=${Math.round(c.rect.width)}x${Math.round(c.rect.height)} scale=${c.scaleX.toFixed(2)} pos=(${Math.round(c.rect.left)},${Math.round(c.rect.top)})`
          )
          .join(' | ')
      )

      // Thử cách 1: đọc canvas trực tiếp
      let outCanvas = this._renderViaCanvases(rects, canvases)

      // Fallback: screenshot
      if (!outCanvas) {
        console.log('[SheetsCapturer] fallback to screenshot')
        this._showToast('Using fallback...')
        outCanvas = await this._renderViaScreenshot(rects)
      }

      if (!outCanvas) {
        this._showToast('Capture failed', true)
        return
      }

      await this._copyToClipboard(outCanvas)
      this._showToast(`Copied! (${outCanvas.width}x${outCanvas.height})`)
    } catch (err) {
      console.error('[SheetsCapturer] error:', err)
      this._showToast('Capture failed', true)
    }
  }
}

window.__SheetsCapturer = new SheetsCapturer()
