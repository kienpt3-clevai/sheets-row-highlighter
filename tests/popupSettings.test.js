const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')

const { getResetSettings, normalizePopupSettings } = require('../scripts/popupSettings.js')

const fallbackDefaults = {
  defaultColor: '#c2185b',
  defaultOpacity: '0.8',
  defaultRow: true,
  defaultColumn: true,
  defaultFillRow: true,
  defaultFillCol: true,
  defaultLineSize: 3.25,
  defaultRowFillOpacity: 0.05,
  defaultColFillOpacity: 0.05,
  defaultRowLineColor: '#c2185b',
  defaultColLineColor: '#c2185b',
  defaultRowFillColor: '#c2185b',
  defaultColFillColor: '#c2185b',
}

test('getResetSettings prefers saved user defaults over hard-coded defaults', () => {
  const userDefaults = {
    color: '#1565c0',
    rowLineColor: '#0e65eb',
    colLineColor: '#1b5e20',
    rowFillColor: '#5c0eec',
    colFillColor: '#ec930e',
    opacity: 0.6,
    row: false,
    column: true,
    fillRow: false,
    fillCol: true,
    lineSize: 1.5,
    rowFillOpacity: 0.15,
    colFillOpacity: 0.2,
  }

  assert.deepEqual(getResetSettings(userDefaults, fallbackDefaults), userDefaults)
})

test('getResetSettings falls back to hard-coded defaults when user defaults are missing', () => {
  assert.deepEqual(getResetSettings(undefined, fallbackDefaults), {
    color: '#c2185b',
    rowLineColor: '#c2185b',
    colLineColor: '#c2185b',
    rowFillColor: '#c2185b',
    colFillColor: '#c2185b',
    opacity: '0.8',
    row: true,
    column: true,
    fillRow: true,
    fillCol: true,
    lineSize: 3.25,
    rowFillOpacity: 0.05,
    colFillOpacity: 0.05,
  })
})

test('normalizePopupSettings merges partial sheet settings over user defaults', () => {
  const sheetSettings = {
    rowLineColor: '#0e65eb',
    fillRow: false,
    rowFillOpacity: 0.25,
  }
  const userDefaults = {
    color: '#1565c0',
    rowLineColor: '#1565c0',
    colLineColor: '#1b5e20',
    rowFillColor: '#5c0eec',
    colFillColor: '#ec930e',
    opacity: 0.6,
    row: false,
    column: true,
    fillRow: true,
    fillCol: false,
    lineSize: 1.5,
    rowFillOpacity: 0.15,
    colFillOpacity: 0.2,
  }

  assert.deepEqual(normalizePopupSettings(sheetSettings, userDefaults, fallbackDefaults), {
    color: '#1565c0',
    rowLineColor: '#0e65eb',
    colLineColor: '#1b5e20',
    rowFillColor: '#5c0eec',
    colFillColor: '#ec930e',
    opacity: 0.6,
    row: false,
    column: true,
    fillRow: false,
    fillCol: false,
    lineSize: 1.5,
    rowFillOpacity: 0.25,
    colFillOpacity: 0.2,
  })
})

test('normalizePopupSettings falls back from cellOpacity when fill opacities are absent', () => {
  assert.deepEqual(
    normalizePopupSettings(
      {
        color: '#6a1b9a',
        cellOpacity: 0.3,
      },
      undefined,
      fallbackDefaults
    ),
    {
      color: '#6a1b9a',
      rowLineColor: '#6a1b9a',
      colLineColor: '#6a1b9a',
      rowFillColor: '#6a1b9a',
      colFillColor: '#6a1b9a',
      opacity: '0.8',
      row: true,
      column: true,
      fillRow: true,
      fillCol: true,
      lineSize: 3.25,
      rowFillOpacity: 0.3,
      colFillOpacity: 0.3,
    }
  )
})

test('spreadsheet content scripts load popupSettings helper before main.js', () => {
  const manifestPath = path.join(__dirname, '..', 'manifest.json')
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
  const spreadsheetScripts = manifest.content_scripts.find((entry) =>
    entry.matches.includes('https://docs.google.com/spreadsheets/d/*')
  )?.js

  assert.ok(Array.isArray(spreadsheetScripts), 'spreadsheet content scripts should exist')
  assert.ok(
    spreadsheetScripts.includes('scripts/popupSettings.js'),
    'popupSettings helper should be injected into spreadsheet pages'
  )
  assert.ok(
    spreadsheetScripts.indexOf('scripts/popupSettings.js') <
      spreadsheetScripts.indexOf('scripts/content/main.js'),
    'popupSettings helper must load before main.js'
  )
})
