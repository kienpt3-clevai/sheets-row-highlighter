// @ts-check

/** Giá trị mặc định dùng chung cho popup, background và content script */
const DEFAULTS = {
  DEFAULT_ROW: true,
  DEFAULT_COLUMN: true,
  DEFAULT_FILL_ROW: true,
  DEFAULT_FILL_COL: true,
  DEFAULT_COLOR: '#c2185b',
  DEFAULT_OPACITY: '1',
  DEFAULT_LINE_SIZE: 3.25,
  DEFAULT_ROW_LINE_SIZE: 3.25,
  DEFAULT_COL_LINE_SIZE: 3.25,
  DEFAULT_CELL_OPACITY: 0.05,
  DEFAULT_ROW_FILL_OPACITY: 0.05,
  DEFAULT_COL_FILL_OPACITY: 0.05,
  // Màu riêng cho line/fill theo Row/Col; mặc định đều dùng DEFAULT_COLOR
  DEFAULT_ROW_LINE_COLOR: '#c2185b',
  DEFAULT_COL_LINE_COLOR: '#c2185b',
  DEFAULT_ROW_FILL_COLOR: '#c2185b',
  DEFAULT_COL_FILL_COLOR: '#c2185b',
}
const g = typeof window !== 'undefined' ? window : globalThis
Object.assign(g, DEFAULTS)

/**
 * @typedef {Object} HighlightRect
 * @property {number} x
 * @property {number} y
 * @property {number} width
 * @property {number} height
 */

/**
 * @typedef {Object} ActiveCellLocator
 * @property {() => Array<HighlightRect>} getHighlightRectList
 * @property {() => Partial<CSSStyleDeclaration>} getSheetContainerStyle
 * @property {() => string} getSheetKey
 * @property {() => HighlightRect | null} [getActiveCellRect]
 */
