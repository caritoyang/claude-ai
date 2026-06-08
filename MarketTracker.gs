// ============================================================
// MARKET TRACKER - Google Apps Script
// Columns: A=Ticker | B=Signal | C=PDH | D=PDL | E=PMH | F=PML | G=Price
//
// 9:30 ET  → fetchDailyLevels()  trae PDH, PDL, PMH, PML (1 vez al día)
// ▶ Run    → runNow()            actualiza Price y Signal
// ============================================================

const SHEET_NAME     = "Tracker";
const COL_TICKER     = 1; // A
const COL_ARROW      = 2; // B
const COL_PDH        = 3; // C
const COL_PDL        = 4; // D
const COL_PMH        = 5; // E
const COL_PML        = 6; // F
const COL_PRICE      = 7; // G

const FINNHUB_API_KEY = "d8jhsl9r01qh6g3qd6pgd8jhsl9r01qh6g3qd6q0";

// ── Precio real-time desde Finnhub ───────────────────────────
function fetchFinnhubPrice(ticker) {
  const url = `https://finnhub.io/api/v1/quote?symbol=${encodeURIComponent(ticker)}&token=${FINNHUB_API_KEY}`;
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      Logger.log(`[${ticker}] quote HTTP ${resp.getResponseCode()}`);
      return null;
    }
    const data  = JSON.parse(resp.getContentText());
    const price = data.c > 0 ? data.c : data.pc;
    return price > 0 ? price : null;
  } catch (e) {
    Logger.log(`[${ticker}] quote error: ${e}`);
    return null;
  }
}

// ── PMH y PML desde Yahoo Finance (pre-market 4:00–9:30 ET) ──
function fetchPreMarketLevels(ticker) {
  const url     = `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1m&range=1d&includePrePost=true`;
  const headers = { "User-Agent": "Mozilla/5.0" };
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers });
    if (resp.getResponseCode() !== 200) return null;

    const result = JSON.parse(resp.getContentText()).chart.result[0];
    const quotes = result.indicators.quote[0];
    const timestamps = result.timestamp;

    const offsetStr  = Utilities.formatDate(new Date(), "America/New_York", "Z");
    const sign       = offsetStr[0] === "-" ? -1 : 1;
    const offsetMins = sign * (parseInt(offsetStr.slice(1,3)) * 60 + parseInt(offsetStr.slice(3,5)));
    const etDateStr  = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");
    const etMidnight = new Date(etDateStr + "T00:00:00Z").getTime() - offsetMins * 60000;
    const preStart   = (etMidnight + 4 * 3600000) / 1000;
    const preEnd     = (etMidnight + 9 * 3600000 + 30 * 60000) / 1000;

    const highs = [], lows = [];
    timestamps.forEach((ts, i) => {
      if (ts >= preStart && ts < preEnd && quotes.high[i] !== null && quotes.low[i] !== null) {
        highs.push(quotes.high[i]);
        lows.push(quotes.low[i]);
      }
    });

    if (highs.length === 0) return null;
    return { pmh: Math.max(...highs), pml: Math.min(...lows) };
  } catch (e) {
    Logger.log(`[${ticker}] Yahoo error: ${e}`);
    return null;
  }
}

// ── Fórmulas GOOGLEFINANCE para PDH y PDL ────────────────────
function setRowFormulas(sheet, row) {
  const a = `A${row}`;
  sheet.getRange(row, COL_PDH)
    .setFormula(`=IFERROR(INDEX(GOOGLEFINANCE(${a},"high",WORKDAY(TODAY(),-1)),2,2),"")`);
  sheet.getRange(row, COL_PDL)
    .setFormula(`=IFERROR(INDEX(GOOGLEFINANCE(${a},"low",WORKDAY(TODAY(),-1)),2,2),"")`);
}

// ── Setup: encabezados + formato ─────────────────────────────
function setupHeaders() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = ["Ticker", "Signal", "PDH", "PDL", "PMH", "PML", "Price"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#263238").setFontColor("#FFFFFF")
    .setFontWeight("bold").setHorizontalAlignment("center");

  sheet.setColumnWidth(COL_TICKER, 90);
  sheet.setColumnWidth(COL_ARROW,  80);
  [COL_PDH, COL_PDL, COL_PMH, COL_PML, COL_PRICE].forEach(c => sheet.setColumnWidth(c, 80));

  const lastRow = Math.max(sheet.getLastRow(), 50);
  [COL_PDH, COL_PDL, COL_PMH, COL_PML, COL_PRICE].forEach(col => {
    sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat("0.00");
  });

  const lastData = sheet.getLastRow();
  for (let row = 2; row <= lastData; row++) {
    if (sheet.getRange(row, COL_TICKER).getValue().toString().trim())
      setRowFormulas(sheet, row);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Setup completo ✓", "Market Tracker", 4);
}

// ── onEdit: agrega fórmulas al escribir un ticker ─────────────
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  if (e.range.getColumn() !== COL_TICKER || e.range.getRow() < 2) return;
  const ticker = e.range.getValue().toString().trim();
  if (ticker) {
    setRowFormulas(sheet, e.range.getRow());
  } else {
    [COL_PDH, COL_PDL, COL_PMH, COL_PML, COL_PRICE, COL_ARROW].forEach(col =>
      sheet.getRange(e.range.getRow(), col).clearContent().setBackground(null));
  }
}

// ── Espera hasta que GOOGLEFINANCE resuelva ───────────────────
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

// ── 9:30 ET: trae PDH, PDL, PMH, PML (1 vez al día) ──────────
function fetchDailyLevels() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  ss.toast("Cargando niveles del día…", "Market Tracker", 60);

  for (let row = 2; row <= lastRow; row++) {
    const ticker = sheet.getRange(row, COL_TICKER).getValue().toString().trim().toUpperCase();
    if (!ticker) continue;

    // PDH/PDL: setear fórmulas y esperar
    setRowFormulas(sheet, row);
    SpreadsheetApp.flush();
    waitForValue(sheet, row, COL_PDH, 12000);
    waitForValue(sheet, row, COL_PDL, 12000);

    // PMH/PML: Yahoo Finance
    const pm = fetchPreMarketLevels(ticker);
    sheet.getRange(row, COL_PMH).setValue(pm ? pm.pmh : "N/A");
    sheet.getRange(row, COL_PML).setValue(pm ? pm.pml : "N/A");

    Utilities.sleep(300);
  }

  ss.toast("Niveles del día cargados ✓", "Market Tracker", 5);
}

// ── ▶ Run: actualiza Price y Signal (rápido) ─────────────────
function runNow() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  ss.toast("Actualizando precios…", "Market Tracker", 30);

  const rowData = [];

  for (let i = 0; i <= lastRow - 2; i++) {
    const sheetRow = i + 2;
    const ticker   = sheet.getRange(sheetRow, COL_TICKER).getValue().toString().trim().toUpperCase();
    if (!ticker) continue;

    // Leer niveles ya cargados
    const pdh = parseFloat(sheet.getRange(sheetRow, COL_PDH).getValue());
    const pdl = parseFloat(sheet.getRange(sheetRow, COL_PDL).getValue());
    const pmh = parseFloat(sheet.getRange(sheetRow, COL_PMH).getValue());
    const pml = parseFloat(sheet.getRange(sheetRow, COL_PML).getValue());

    // Obtener precio actual
    const price = fetchFinnhubPrice(ticker);
    if (price !== null) sheet.getRange(sheetRow, COL_PRICE).setValue(price);

    rowData.push({ sheetRow, ticker, pdh, pdl, pmh, pml, price });
    Utilities.sleep(200);
  }

  // Evaluar señales
  let updated = 0;
  rowData.forEach(({ sheetRow, ticker, pdh, pdl, pmh, pml, price }) => {
    const arrowCell = sheet.getRange(sheetRow, COL_ARROW);

    if (isNaN(pdh) || isNaN(pdl) || price === null) {
      arrowCell.setValue("N/A").setFontColor("#F44336").setFontSize(12)
               .setHorizontalAlignment("center").setBackground(null);
      Logger.log(`${ticker}: faltan niveles — corré "Cargar niveles del día" primero`);
      return;
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
  });

  ss.toast(`${updated} ticker(s) actualizados ✓`, "Market Tracker", 5);
}

// ── Triggers: 9:30 para niveles, manual para precio ──────────
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["fetchDailyLevels", "runNow"].includes(t.getHandlerFunction()))
      ScriptApp.deleteTrigger(t);
  });

  const days = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY,
                ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY,
                ScriptApp.WeekDay.FRIDAY];

  days.forEach(day => {
    ScriptApp.newTrigger("fetchDailyLevels")
      .timeBased().atHour(9).nearMinute(30).everyWeeks(1)
      .onWeekDay(day).inTimezone("America/New_York").create();
  });

  SpreadsheetApp.getActiveSpreadsheet()
    .toast("Trigger instalado: fetchDailyLevels Lun-Vie 9:30 ET ✓", "Market Tracker", 6);
}

// ── Menú ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📈 Market Tracker")
    .addItem("1. Setup headers y fórmulas", "setupHeaders")
    .addItem("2. Instalar trigger (9:30 ET, Lun-Vie)", "installTriggers")
    .addSeparator()
    .addItem("📥 Cargar niveles del día (PDH/PDL/PMH/PML)", "fetchDailyLevels")
    .addItem("▶ Run (actualiza Price + Signal)", "runNow")
    .addToUi();
}
