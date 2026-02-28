// @ts-check

/** Giá trị mặc định dùng chung cho popup, background và content script */
const DEFAULTS = {
  DEFAULT_ROW: true,
  DEFAULT_COLUMN: true,
  DEFAULT_COLOR: '#c2185b',
  DEFAULT_OPACITY: '0.8',
  DEFAULT_LINE_SIZE: 3.25,
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
 */
