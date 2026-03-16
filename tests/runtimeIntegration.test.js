const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const vm = require('node:vm')

const popupScript = fs.readFileSync(
  path.join(__dirname, '..', 'scripts', 'popup.js'),
  'utf8'
)
const mainScript = fs.readFileSync(
  path.join(__dirname, '..', 'scripts', 'content', 'main.js'),
  'utf8'
)
const popupSettingsUtils = require('../scripts/popupSettings.js')

const flush = async () => {
  await new Promise((resolve) => setImmediate(resolve))
  await new Promise((resolve) => setImmediate(resolve))
}

const createElement = (id, initial = {}) => {
  const listeners = {}
  return {
    id,
    value: '',
    checked: false,
    style: {},
    children: [],
    appendChild(child) {
      this.children.push(child)
    },
    addEventListener(type, handler) {
      listeners[type] = listeners[type] || []
      listeners[type].push(handler)
    },
    dispatch(type, event = {}) {
      for (const handler of listeners[type] || []) {
        handler({ target: this, ...event })
      }
    },
    click() {
      this.dispatch('click')
    },
    ...initial,
  }
}

const createPopupHarness = ({ sheetSettings = {}, defaultSettings = {}, sheetKey = 'sheet-1' } = {}) => {
  const elements = new Map()
  const ids = [
    'row',
    'column',
    'rowLineSize',
    'colLineSize',
    'rowOpacity',
    'colOpacity',
    'fillRow',
    'fillCol',
    'reset',
    'setDefault',
    'setDefaultAll',
    'prevRowLineColor',
    'nextRowLineColor',
    'prevColLineColor',
    'nextColLineColor',
    'prevRowFillColor',
    'nextRowFillColor',
    'prevColFillColor',
    'nextColFillColor',
    'rowLineColor',
    'colLineColor',
    'rowFillColor',
    'colFillColor',
  ]

  for (const id of ids) {
    elements.set(id, createElement(id))
  }

  const windowListeners = {}
  const storageState = {
    sheetSettings: { ...sheetSettings },
    defaultSettings: { ...defaultSettings },
  }
  const storageSets = []
  const sentMessages = []
  const huebees = []

  class FakeHuebee {
    constructor(selector) {
      this.selector = selector
      this.color = null
      this.handlers = { change: new Set() }
      huebees.push(this)
    }

    setColor(color) {
      this.color = color
      for (const handler of this.handlers.change) {
        handler(color)
      }
    }

    on(type, handler) {
      this.handlers[type] = this.handlers[type] || new Set()
      this.handlers[type].add(handler)
    }

    off(type, handler) {
      this.handlers[type]?.delete(handler)
    }
  }

  const document = {
    body: createElement('body'),
    getElementById(id) {
      return elements.get(id)
    },
  }

  const chrome = {
    tabs: {
      async query() {
        return [{ id: 1 }]
      },
      async sendMessage(_tabId, message) {
        sentMessages.push(message)
        if (message.type === 'getSheetKey') {
          return { sheetKey }
        }
        return {}
      },
    },
    storage: {
      local: {
        get(keys, callback) {
          const result = {}
          for (const key of keys) {
            result[key] = storageState[key]
          }
          callback(result)
        },
        set(value, callback) {
          Object.assign(storageState, value)
          storageSets.push(value)
          if (callback) callback()
        },
      },
    },
    runtime: {
      onMessage: {
        addListener() {},
      },
    },
  }

  const context = vm.createContext({
    console,
    setTimeout,
    clearTimeout,
    window: {
      addEventListener(type, handler) {
        windowListeners[type] = handler
      },
    },
    document,
    chrome,
    Huebee: FakeHuebee,
    PopupSettingsUtils: popupSettingsUtils,
    DEFAULT_COLOR: '#c2185b',
    DEFAULT_OPACITY: '1',
    DEFAULT_ROW: true,
    DEFAULT_COLUMN: true,
    DEFAULT_FILL_ROW: true,
    DEFAULT_FILL_COL: true,
    DEFAULT_LINE_SIZE: 3.25,
    DEFAULT_ROW_FILL_OPACITY: 0.05,
    DEFAULT_COL_FILL_OPACITY: 0.05,
    DEFAULT_ROW_LINE_COLOR: '#c2185b',
    DEFAULT_COL_LINE_COLOR: '#c2185b',
    DEFAULT_ROW_FILL_COLOR: '#c2185b',
    DEFAULT_COL_FILL_COLOR: '#c2185b',
  })

  vm.runInContext(popupScript, context, { filename: 'popup.js' })

  return {
    elements,
    storageState,
    storageSets,
    sentMessages,
    async triggerLoad() {
      await windowListeners.load()
      await flush()
    },
  }
}

test('popup initialization and reset both honor saved defaultSettings', async () => {
  const harness = createPopupHarness({
    sheetSettings: {
      'sheet-1': {
        color: '#ec0e2f',
        rowLineColor: '#ec0e2f',
        colLineColor: '#ec0e2f',
        rowFillColor: '#ec0e2f',
        colFillColor: '#ec0e2f',
        row: true,
        column: false,
        fillRow: true,
        fillCol: false,
        rowLineSize: 4,
        colLineSize: 2,
        rowFillOpacity: 0.1,
        colFillOpacity: 0.15,
      },
    },
    defaultSettings: {
      color: '#1565c0',
      rowLineColor: '#0e65eb',
      colLineColor: '#1b5e20',
      rowFillColor: '#5c0eec',
      colFillColor: '#ec930e',
      row: false,
      column: true,
      fillRow: false,
      fillCol: true,
      rowLineSize: 1.5,
      colLineSize: 2.5,
      rowFillOpacity: 0.15,
      colFillOpacity: 0.2,
    },
  })

  await harness.triggerLoad()

  assert.equal(harness.elements.get('row').checked, true)
  assert.equal(harness.elements.get('column').checked, false)
  assert.equal(harness.elements.get('rowLineSize').value, 4)
  assert.equal(harness.elements.get('colLineSize').value, 2)

  harness.elements.get('reset').click()
  await flush()

  assert.equal(harness.elements.get('row').checked, false)
  assert.equal(harness.elements.get('column').checked, true)
  assert.equal(harness.elements.get('fillRow').checked, false)
  assert.equal(harness.elements.get('fillCol').checked, true)
  assert.equal(harness.elements.get('rowLineSize').value, 1.5)
  assert.equal(harness.elements.get('colLineSize').value, 2.5)
  assert.equal(harness.elements.get('rowOpacity').value, '0.15')
  assert.equal(harness.elements.get('colOpacity').value, '0.2')
  assert.deepEqual({ ...harness.storageState.sheetSettings['sheet-1'] }, {
    color: '#0e65eb',
    rowLineColor: '#0e65eb',
    colLineColor: '#1b5e20',
    rowFillColor: '#5c0eec',
    colFillColor: '#ec930e',
    row: false,
    column: true,
    fillRow: false,
    fillCol: true,
    opacity: '1',
    rowLineSize: 1.5,
    colLineSize: 2.5,
    rowFillOpacity: 0.15,
    colFillOpacity: 0.2,
    cellOpacity: 0.15,
    updatedAt: harness.storageState.sheetSettings['sheet-1'].updatedAt,
  })
})

test('content script applies saved defaultSettings when sheet has no per-sheet config', async () => {
  const appInstances = []
  const storageListeners = []

  class FakeRowHighlighterApp {
    constructor(appContainer, locator) {
      this.appContainer = appContainer
      this.locator = locator
      this.backgroundColor = '#0e65eb'
      this.rowLineColor = '#0e65eb'
      this.colLineColor = '#0e65eb'
      this.rowFillColor = '#0e65eb'
      this.colFillColor = '#0e65eb'
      this.opacity = '1'
      this.rowLineSize = 3.25
      this.colLineSize = 3.25
      this.rowFillOpacity = 0.05
      this.colFillOpacity = 0.05
      this.isRowEnabled = true
      this.isColEnabled = false
      this.fillRowEnabled = true
      this.fillColEnabled = true
      this.updateCalls = 0
      appInstances.push(this)
    }

    update() {
      this.updateCalls += 1
    }
  }

  class FakeLocator {
    getSheetKey() {
      return 'sheet-1'
    }

    getHighlightRectList() {
      return []
    }

    getSheetContainerStyle() {
      return {}
    }
  }

  const storageState = {
    sheetSettings: {},
    defaultSettings: {
      color: '#1565c0',
      rowLineColor: '#0e65eb',
      colLineColor: '#1b5e20',
      rowFillColor: '#5c0eec',
      colFillColor: '#ec930e',
      row: false,
      column: true,
      fillRow: false,
      fillCol: true,
      rowLineSize: 1.5,
      colLineSize: 2.5,
      rowFillOpacity: 0.15,
      colFillOpacity: 0.2,
    },
  }

  const document = {
    body: createElement('body'),
    addEventListener() {},
    createElement(tag) {
      return createElement(tag)
    },
  }

  const runtimeListeners = []
  const chrome = {
    storage: {
      local: {
        get(keys, callback) {
          const result = {}
          for (const key of keys) {
            result[key] = storageState[key]
          }
          callback(result)
        },
        set() {},
      },
      onChanged: {
        addListener(listener) {
          storageListeners.push(listener)
        },
      },
    },
    runtime: {
      onMessage: {
        addListener(listener) {
          runtimeListeners.push(listener)
        },
      },
    },
  }

  const context = vm.createContext({
    console,
    setTimeout,
    clearTimeout,
    window: {
      top: null,
      addEventListener() {},
    },
    document,
    chrome,
    PopupSettingsUtils: popupSettingsUtils,
    SheetsActiveCellLocator: FakeLocator,
    RowHighlighterApp: FakeRowHighlighterApp,
    DEFAULT_COLOR: '#c2185b',
    DEFAULT_OPACITY: '1',
    DEFAULT_ROW: true,
    DEFAULT_COLUMN: true,
    DEFAULT_FILL_ROW: true,
    DEFAULT_FILL_COL: true,
    DEFAULT_LINE_SIZE: 3.25,
    DEFAULT_ROW_FILL_OPACITY: 0.05,
    DEFAULT_COL_FILL_OPACITY: 0.05,
    DEFAULT_ROW_LINE_COLOR: '#c2185b',
    DEFAULT_COL_LINE_COLOR: '#c2185b',
    DEFAULT_ROW_FILL_COLOR: '#c2185b',
    DEFAULT_COL_FILL_COLOR: '#c2185b',
  })
  context.window.top = context.window

  vm.runInContext(mainScript, context, { filename: 'main.js' })
  await flush()

  const app = appInstances[0]
  assert.ok(app, 'content script should create a RowHighlighterApp instance')
  assert.equal(app.backgroundColor, '#1565c0')
  assert.equal(app.rowLineColor, '#0e65eb')
  assert.equal(app.colLineColor, '#1b5e20')
  assert.equal(app.rowFillColor, '#5c0eec')
  assert.equal(app.colFillColor, '#ec930e')
  assert.equal(app.opacity, '1')
  assert.equal(app.isRowEnabled, false)
  assert.equal(app.isColEnabled, true)
  assert.equal(app.fillRowEnabled, false)
  assert.equal(app.fillColEnabled, true)
  assert.equal(app.rowLineSize, 1.5)
  assert.equal(app.colLineSize, 2.5)
  assert.equal(app.rowFillOpacity, 0.15)
  assert.equal(app.colFillOpacity, 0.2)
  assert.equal(app.updateCalls, 1)
  assert.equal(storageListeners.length, 1)
  assert.equal(runtimeListeners.length, 1)
})
