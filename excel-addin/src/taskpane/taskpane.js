/* global Excel, Office */

// ============================================================
// Excel Row Highlighter - Office Add-in (JS)
// Highlights active row/column using cell formatting + CF
// ============================================================

// --- State ---
let highlightEnabled = true;
let prevRow = -1;
let prevCol = -1;
let prevRowCount = 0;
let prevColCount = 0;

// --- Settings (defaults) ---
let settings = {
  rowLineEnabled: true,
  colLineEnabled: true,
  rowFillEnabled: true,
  rowLineColor: "#c2185b",
  colLineColor: "#3399ff",
  rowFillColor: "#c2185b",
  rowFillOpacity: 0.15,
  rowLineWeight: "Medium", // Thin, Medium, Thick
  colLineWeight: "Medium",
};

// --- Helpers ---
function blendColor(hex, opacity) {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const br = Math.round(255 - (255 - r) * opacity);
  const bg = Math.round(255 - (255 - g) * opacity);
  const bb = Math.round(255 - (255 - b) * opacity);
  return (
    "#" +
    br.toString(16).padStart(2, "0") +
    bg.toString(16).padStart(2, "0") +
    bb.toString(16).padStart(2, "0")
  );
}

function colLetter(colNum) {
  let s = "";
  while (colNum > 0) {
    colNum--;
    s = String.fromCharCode((colNum % 26) + 65) + s;
    colNum = Math.floor(colNum / 26);
  }
  return s;
}

function loadSettings() {
  try {
    const saved = localStorage.getItem("rh_settings");
    if (saved) {
      Object.assign(settings, JSON.parse(saved));
    }
  } catch (e) {
    // ignore
  }
}

function saveSettings() {
  try {
    localStorage.setItem("rh_settings", JSON.stringify(settings));
  } catch (e) {
    // ignore
  }
}

// --- Clear previous highlights ---
async function clearHighlights(context) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();

  // Remove ALL conditional formats (our CF rules)
  try {
    sheet.conditionalFormats.clearAll();
  } catch (e) {
    // ignore if not supported
  }

  // Clear previous row borders
  if (prevRow > 0) {
    try {
      const cols = Math.max(prevCol + 30, 50);
      const prevRowRange = sheet.getRangeByIndexes(
        prevRow - 1, 0, prevRowCount || 1, cols
      );
      prevRowRange.format.borders.getItem("EdgeTop").style = "None";
      prevRowRange.format.borders.getItem("EdgeBottom").style = "None";
    } catch (e) {
      // ignore
    }
  }

  // Clear previous col borders
  if (prevCol > 0) {
    try {
      const rows = Math.max(prevRow + 50, 100);
      const prevColRange = sheet.getRangeByIndexes(
        0, prevCol - 1, rows, prevColCount || 1
      );
      prevColRange.format.borders.getItem("EdgeLeft").style = "None";
      prevColRange.format.borders.getItem("EdgeRight").style = "None";
    } catch (e) {
      // ignore
    }
  }
}

// --- Draw highlights ---
async function drawHighlights(context) {
  if (!highlightEnabled) return;

  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const activeCell = context.workbook.getActiveCell();
  const selection = context.workbook.getSelectedRange();

  activeCell.load("rowIndex, columnIndex");
  selection.load("rowIndex, columnIndex, rowCount, columnCount");
  await context.sync();

  const row = selection.rowIndex + 1; // 1-based
  const col = selection.columnIndex + 1;
  const rowCount = selection.rowCount;
  const colCount = selection.columnCount;

  // Skip if nothing changed
  if (
    row === prevRow &&
    col === prevCol &&
    rowCount === prevRowCount &&
    colCount === prevColCount
  ) {
    return;
  }

  // Clear old highlights
  await clearHighlights(context);
  await context.sync();

  // --- Row Fill (CF - works in freeze panes) ---
  if (settings.rowFillEnabled && settings.rowFillOpacity > 0) {
    const fillColor = blendColor(settings.rowFillColor, settings.rowFillOpacity);
    const rowStart = row;
    const rowEnd = row + rowCount - 1;

    const cfRange = sheet.getRange("1:1048576"); // entire sheet
    const formula = `=AND(ROW()>=${rowStart},ROW()<=${rowEnd})`;
    const cf = cfRange.conditionalFormats.add(
      Excel.ConditionalFormatType.custom
    );
    cf.custom.rule.formula = formula;
    cf.custom.format.fill.color = fillColor;
    cf.stopIfTrue = false;
  }

  // Get visible range extent for lines
  const visRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1").getResizedRange(99, 49);
  // Use a generous fixed range: 100 rows x 50 cols from origin, or based on active area
  const activeCell2 = context.workbook.getActiveCell();
  activeCell2.load("rowIndex, columnIndex");
  await context.sync();

  // Extend lines well beyond visible area
  const lineRows = Math.max(activeCell2.rowIndex + 50, 100);
  const lineCols = Math.max(activeCell2.columnIndex + 30, 50);

  // --- Row Lines (borders on the row range) ---
  if (settings.rowLineEnabled) {
    const rowRange = sheet.getRangeByIndexes(row - 1, 0, rowCount, lineCols);

    const topBorder = rowRange.format.borders.getItem("EdgeTop");
    topBorder.style = "Continuous";
    topBorder.color = settings.rowLineColor;
    topBorder.weight = settings.rowLineWeight;

    const bottomBorder = rowRange.format.borders.getItem("EdgeBottom");
    bottomBorder.style = "Continuous";
    bottomBorder.color = settings.rowLineColor;
    bottomBorder.weight = settings.rowLineWeight;
  }

  // --- Col Lines (borders on the column range) ---
  if (settings.colLineEnabled) {
    const colRange = sheet.getRangeByIndexes(0, col - 1, lineRows, colCount);

    const leftBorder = colRange.format.borders.getItem("EdgeLeft");
    leftBorder.style = "Continuous";
    leftBorder.color = settings.colLineColor;
    leftBorder.weight = settings.colLineWeight;

    const rightBorder = colRange.format.borders.getItem("EdgeRight");
    rightBorder.style = "Continuous";
    rightBorder.color = settings.colLineColor;
    rightBorder.weight = settings.colLineWeight;
  }

  // Update tracking
  prevRow = row;
  prevCol = col;
  prevRowCount = rowCount;
  prevColCount = colCount;
}

// --- Selection change handler ---
async function onSelectionChanged() {
  const dbg = document.getElementById("status");
  try {
    await Excel.run(async (context) => {
      if (dbg) dbg.textContent = "Drawing...";
      await drawHighlights(context);
      await context.sync();
      if (dbg) dbg.textContent = highlightEnabled ? "ON" : "OFF";
    });
  } catch (error) {
    if (dbg) dbg.textContent = "ERR: " + error.message;
    console.log("RH selection change error:", error);
  }
}

// --- Toggle highlight ---
async function toggleHighlightState() {
  highlightEnabled = !highlightEnabled;
  if (!highlightEnabled) {
    try {
      await Excel.run(async (context) => {
        await clearHighlights(context);
        await context.sync();
      });
    } catch (e) {
      // ignore
    }
    prevRow = -1;
    prevCol = -1;
  } else {
    await onSelectionChanged();
  }
  updateUI();
}

// --- Register event handlers ---
let selectionHandler = null;

async function registerEvents() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      selectionHandler = sheet.onSelectionChanged.add(onSelectionChanged);

      // Re-register on sheet change (when user switches sheets)
      context.workbook.worksheets.onActivated.add(async () => {
        await Excel.run(async (ctx) => {
          const newSheet = ctx.workbook.worksheets.getActiveWorksheet();
          newSheet.onSelectionChanged.add(onSelectionChanged);
          await ctx.sync();
        });
        // Redraw on new sheet
        await onSelectionChanged();
      });

      await context.sync();
    });
  } catch (error) {
    console.log("RH register events error:", error);
  }
}

// --- Save cleanup (skip if events not available) ---
async function registerSaveCleanup() {
  // onWillSave/onSaved not available in all Excel versions — skip silently
}

// --- UI Logic ---
function updateUI() {
  const statusEl = document.getElementById("status");
  if (statusEl) {
    statusEl.textContent = highlightEnabled ? "ON" : "OFF";
    statusEl.className = highlightEnabled ? "status-on" : "status-off";
  }

  // Sync checkboxes
  const chkRowLine = document.getElementById("chkRowLine");
  const chkColLine = document.getElementById("chkColLine");
  const chkRowFill = document.getElementById("chkRowFill");
  if (chkRowLine) chkRowLine.checked = settings.rowLineEnabled;
  if (chkColLine) chkColLine.checked = settings.colLineEnabled;
  if (chkRowFill) chkRowFill.checked = settings.rowFillEnabled;

  const rowLineColor = document.getElementById("rowLineColor");
  const colLineColor = document.getElementById("colLineColor");
  const rowFillColor = document.getElementById("rowFillColor");
  const rowFillOpacity = document.getElementById("rowFillOpacity");
  if (rowLineColor) rowLineColor.value = settings.rowLineColor;
  if (colLineColor) colLineColor.value = settings.colLineColor;
  if (rowFillColor) rowFillColor.value = settings.rowFillColor;
  if (rowFillOpacity) rowFillOpacity.value = settings.rowFillOpacity;
}

function setupUI() {
  // Toggle button
  const btnToggle = document.getElementById("btnToggle");
  if (btnToggle) {
    btnToggle.addEventListener("click", toggleHighlightState);
  }

  // Checkboxes
  document.getElementById("chkRowLine")?.addEventListener("change", (e) => {
    settings.rowLineEnabled = e.target.checked;
    saveSettings();
  });
  document.getElementById("chkColLine")?.addEventListener("change", (e) => {
    settings.colLineEnabled = e.target.checked;
    saveSettings();
  });
  document.getElementById("chkRowFill")?.addEventListener("change", (e) => {
    settings.rowFillEnabled = e.target.checked;
    saveSettings();
  });

  // Colors
  document.getElementById("rowLineColor")?.addEventListener("input", (e) => {
    settings.rowLineColor = e.target.value;
    saveSettings();
  });
  document.getElementById("colLineColor")?.addEventListener("input", (e) => {
    settings.colLineColor = e.target.value;
    saveSettings();
  });
  document.getElementById("rowFillColor")?.addEventListener("input", (e) => {
    settings.rowFillColor = e.target.value;
    saveSettings();
  });

  // Opacity
  document
    .getElementById("rowFillOpacity")
    ?.addEventListener("input", (e) => {
      settings.rowFillOpacity = parseFloat(e.target.value);
      document.getElementById("opacityValue").textContent =
        settings.rowFillOpacity.toFixed(2);
      saveSettings();
    });

  updateUI();
}

// --- Init ---
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    loadSettings();
    setupUI();
    await registerEvents();
    await registerSaveCleanup();

    // Initial highlight
    if (highlightEnabled) {
      await onSelectionChanged();
    }
  }
});

// Export for commands
window.toggleHighlight = toggleHighlightState;
