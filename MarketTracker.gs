// ============================================================
// MARKET TRACKER - Google Apps Script
// Columns: A=Ticker | B=Signal | C=PDH | D=PDL | E=PMH | F=PML | G=Price
//
// PDH, PDL  → GOOGLEFINANCE (día anterior, delay OK)
// Price     → Finnhub real-time
// PMH, PML  → ingreso manual cada mañana
// ============================================================

const SHEET_NAME     = "Tracker";
const COL_TICKER     = 1; // A
const COL_ARROW      = 2; // B
const COL_PDH        = 3; // C
const COL_PDL        = 4; // D
const COL_PMH        = 5; // E  ← manual
const COL_PML        = 6; // F  ← manual
const COL_PRICE      = 7; // G  ← Finnhub real-time

// ⚠️ Pegá tu API key de finnhub.io acá:
const FINNHUB_API_KEY = "TU_API_KEY_ACÁ";

// ── Precio real-time desde Finnhub ───────────────────────────
function fetchFinnhubPrice(ticker) {
  const url = `https://finnhub.io/api/v1/quote?symbol=${encodeURIComponent(ticker)}&token=${FINNHUB_API_KEY}`;
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      Logger.log(`[${ticker}] Finnhub HTTP ${resp.getResponseCode()}: ${resp.getContentText().substring(0, 200)}`);
      return null;
    }
    const data = JSON.parse(resp.getContentText());
    // c = current price (0 si el mercado está cerrado, usar pc = previous close)
    const price = data.c > 0 ? data.c : data.pc;
    Logger.log(`[${ticker}] Finnhub price=${price}`);
    return price > 0 ? price : null;
  } catch (e) {
    Logger.log(`[${ticker}] Finnhub error: ${e}`);
    return null;
  }
}

// ── Fórmulas GOOGLEFINANCE para PDH y PDL (día anterior) ─────
function setRowFormulas(sheet, row) {
  const a = `A${row}`;
  sheet.getRange(row, COL_PDH)
    .setFormula(`=IFERROR(INDEX(GOOGLEFINANCE(${a},"high",WORKDAY(TODAY(),-1)),2,2),"")`);
  sheet.getRange(row, COL_PDL)
    .setFormula(`=IFERROR(INDEX(GOOGLEFINANCE(${a},"low",WORKDAY(TODAY(),-1)),2,2),"")`);
}

// ── Setup: encabezados + fórmulas para todos los tickers ─────
function setupHeaders() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const headers = ["Ticker", "Signal", "PDH", "PDL", "PMH ✍", "PML ✍", "Price"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#263238")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setColumnWidth(COL_TICKER, 90);
  sheet.setColumnWidth(COL_ARROW,  80);
  sheet.setColumnWidth(COL_PDH,    80);
  sheet.setColumnWidth(COL_PDL,    80);
  sheet.setColumnWidth(COL_PMH,    90);
  sheet.setColumnWidth(COL_PML,    90);
  sheet.setColumnWidth(COL_PRICE,  80);

  // Columnas manuales en amarillo
  const lastRow = Math.max(sheet.getLastRow(), 10);
  sheet.getRange(2, COL_PMH, lastRow - 1, 1).setBackground("#FFFDE7");
  sheet.getRange(2, COL_PML, lastRow - 1, 1).setBackground("#FFFDE7");

  // Fórmulas PDH/PDL para filas existentes
  const lastData = sheet.getLastRow();
  for (let row = 2; row <= lastData; row++) {
    const ticker = sheet.getRange(row, COL_TICKER).getValue().toString().trim();
    if (ticker) setRowFormulas(sheet, row);
  }

  ss.toast("Setup completo ✓", "Market Tracker", 4);
}

// ── onEdit: agrega fórmulas PDH/PDL al escribir un ticker ────
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  if (e.range.getColumn() !== COL_TICKER) return;
  const row = e.range.getRow();
  if (row < 2) return;

  const ticker = e.range.getValue().toString().trim();
  if (ticker) {
    setRowFormulas(sheet, row);
  } else {
    sheet.getRange(row, COL_PDH).clearContent();
    sheet.getRange(row, COL_PDL).clearContent();
    sheet.getRange(row, COL_PRICE).clearContent();
    sheet.getRange(row, COL_ARROW).clearContent().setBackground(null);
  }
}

// ── Espera hasta que GOOGLEFINANCE resuelva en una celda ──────
function waitForValue(sheet, row, col, maxWaitMs) {
  const deadline = Date.now() + maxWaitMs;
  while (Date.now() < deadline) {
    SpreadsheetApp.flush();
    const val = parseFloat(sheet.getRange(row, col).getValue());
    if (!isNaN(val) && val > 0) return val;
    Utilities.sleep(1500);
  }
  return null;
}

// ── Run: obtiene precio real-time y evalúa señales ───────────
function runNow() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (FINNHUB_API_KEY === "TU_API_KEY_ACÁ") {
    SpreadsheetApp.getUi().alert("⚠️ Antes de continuar, pegá tu API key de Finnhub en la línea:\nconst FINNHUB_API_KEY = \"TU_API_KEY_ACÁ\"");
    return;
  }

  if (lastRow < 2) {
    ss.toast("No hay tickers en la columna A.", "Market Tracker", 5);
    return;
  }

  // Asegurar fórmulas PDH/PDL en todas las filas
  for (let row = 2; row <= lastRow; row++) {
    const ticker = sheet.getRange(row, COL_TICKER).getValue().toString().trim();
    if (ticker) setRowFormulas(sheet, row);
  }

  ss.toast("Cargando datos…", "Market Tracker", 30);
  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  let updated = 0;

  for (let i = 0; i <= lastRow - 2; i++) {
    const sheetRow = i + 2;
    const ticker   = sheet.getRange(sheetRow, COL_TICKER).getValue().toString().trim().toUpperCase();
    if (!ticker) continue;

    const arrowCell = sheet.getRange(sheetRow, COL_ARROW);
    arrowCell.setValue("⏳").setFontColor("#9E9E9E").setFontSize(14)
             .setHorizontalAlignment("center").setBackground(null);
    SpreadsheetApp.flush();

    // PDH/PDL desde GOOGLEFINANCE (delay no importa, son datos del día anterior)
    const pdh = waitForValue(sheet, sheetRow, COL_PDH, 12000);
    const pdl = waitForValue(sheet, sheetRow, COL_PDL, 12000);

    // Precio real-time desde Finnhub
    const price = fetchFinnhubPrice(ticker);
    if (price !== null) {
      sheet.getRange(sheetRow, COL_PRICE).setValue(price);
    }

    const pmh = parseFloat(sheet.getRange(sheetRow, COL_PMH).getValue());
    const pml = parseFloat(sheet.getRange(sheetRow, COL_PML).getValue());

    if (pdh === null || pdl === null || price === null) {
      arrowCell.setValue("N/A").setFontColor("#F44336").setFontSize(12)
               .setHorizontalAlignment("center").setBackground(null);
      Logger.log(`${ticker}: sin datos — pdh=${pdh} pdl=${pdl} price=${price}`);
      continue;
    }

    Logger.log(`${ticker}: pdh=${pdh} pdl=${pdl} pmh=${pmh} pml=${pml} price=${price}`);

    const isGreen = !isNaN(pmh) ? (price > pdh && price > pmh) : (price > pdh);
    const isRed   = !isNaN(pml) ? (price < pdl && price < pml) : (price < pdl);

    if (isGreen) {
      arrowCell.setValue("▲").setFontColor("#00C853").setFontSize(16)
               .setHorizontalAlignment("center").setBackground("#E8F5E9");
    } else if (isRed) {
      arrowCell.setValue("▼").setFontColor("#D50000").setFontSize(16)
               .setHorizontalAlignment("center").setBackground("#FFEBEE");
    } else {
      arrowCell.setValue("—").setFontColor("#9E9E9E").setFontSize(16)
               .setHorizontalAlignment("center").setBackground(null);
    }

    updated++;
    Utilities.sleep(200); // respetar límite de 60 req/min de Finnhub
  }

  ss.toast(`${updated} ticker(s) actualizados ✓`, "Market Tracker", 5);
}

// ── Trigger automático a las 9:31 ET ─────────────────────────
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["runNow", "onEdit"].includes(t.getHandlerFunction()))
      ScriptApp.deleteTrigger(t);
  });

  [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY,
   ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY].forEach(day => {
    ScriptApp.newTrigger("runNow")
      .timeBased().atHour(9).nearMinute(31).everyWeeks(1)
      .onWeekDay(day).inTimezone("America/New_York").create();
  });

  SpreadsheetApp.getActiveSpreadsheet()
    .toast("Trigger instalado: Lun-Vie 9:31 ET ✓", "Market Tracker", 6);
}

// ── Menú ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📈 Market Tracker")
    .addItem("1. Setup headers y fórmulas", "setupHeaders")
    .addItem("2. Instalar trigger (9:31 ET, Lun-Vie)", "installTriggers")
    .addSeparator()
    .addItem("▶ Run", "runNow")
    .addToUi();
}
