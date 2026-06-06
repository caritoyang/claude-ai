// ============================================================
// MARKET TRACKER - Google Apps Script
// Columns: A=Ticker | B=Arrow | C=PDH | D=PDL | E=PMH | F=PML
// ============================================================

const SHEET_NAME = "Tracker";
const COL_TICKER = 1; // A
const COL_ARROW  = 2; // B
const COL_PDH    = 3; // C
const COL_PDL    = 4; // D
const COL_PMH    = 5; // E
const COL_PML    = 6; // F

const NEAR_THRESHOLD = 5; // points within which is considered "near"

// ── Headers ──────────────────────────────────────────────────
function setupHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = ["Ticker", "Signal", "PDH", "PDL", "PMH", "PML"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style header row
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#263238")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Set column widths
  sheet.setColumnWidth(COL_TICKER, 90);
  sheet.setColumnWidth(COL_ARROW,  80);
  [COL_PDH, COL_PDL, COL_PMH, COL_PML].forEach(c => sheet.setColumnWidth(c, 80));
}

// ── Fetch data from Yahoo Finance ────────────────────────────
function fetchYahooData(ticker) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1d&range=5d&includePrePost=true`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    const result = json.chart.result[0];
    const meta   = result.meta;

    // Previous day OHLC (index -2 from today, i.e. last completed session)
    const timestamps = result.timestamp;
    const quotes     = result.indicators.quote[0];
    const today      = new Date();
    today.setHours(0, 0, 0, 0);

    // Find the index of the last completed regular session (yesterday or last trading day)
    let prevIdx = -1;
    for (let i = timestamps.length - 1; i >= 0; i--) {
      const d = new Date(timestamps[i] * 1000);
      d.setHours(0, 0, 0, 0);
      if (d < today) { prevIdx = i; break; }
    }

    const pdh = prevIdx >= 0 ? quotes.high[prevIdx]  : null;
    const pdl = prevIdx >= 0 ? quotes.low[prevIdx]   : null;

    // Pre-market high/low from extended hours (preMarketPrice not in chart v8 directly)
    // Use the summary detail endpoint for pre-market price snapshot
    const summaryUrl = `https://query1.finance.yahoo.com/v11/finance/quoteSummary/${encodeURIComponent(ticker)}?modules=summaryDetail,price`;
    const summaryResp = UrlFetchApp.fetch(summaryUrl, { muteHttpExceptions: true });
    const summaryJson = JSON.parse(summaryResp.getContentText());
    const priceData   = summaryJson.quoteSummary.result[0].price;

    // Yahoo provides pre-market high/low via the quoteType + price module
    const pmh = priceData.preMarketHigh  ? priceData.preMarketHigh.raw  : null;
    const pml = priceData.preMarketLow   ? priceData.preMarketLow.raw   : null;

    // Current / pre-market price (best available at call time)
    const currentPrice = priceData.preMarketPrice && priceData.preMarketPrice.raw
      ? priceData.preMarketPrice.raw
      : (priceData.regularMarketPrice ? priceData.regularMarketPrice.raw : null);

    return { pdh, pdl, pmh, pml, currentPrice };
  } catch (e) {
    Logger.log(`Error fetching ${ticker}: ${e}`);
    return null;
  }
}

// ── Run — fetch current data and evaluate signals ─────────────
function runNow() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const tickers = sheet.getRange(2, COL_TICKER, lastRow - 1, 1).getValues();

  tickers.forEach((row, i) => {
    const ticker = row[0].toString().trim().toUpperCase();
    if (!ticker) return;

    const data = fetchYahooData(ticker);
    if (!data) return;

    const sheetRow  = i + 2;
    const { pdh, pdl, pmh, pml, currentPrice } = data;

    // Populate level columns
    sheet.getRange(sheetRow, COL_PDH).setValue(pdh !== null ? pdh : "N/A");
    sheet.getRange(sheetRow, COL_PDL).setValue(pdl !== null ? pdl : "N/A");
    sheet.getRange(sheetRow, COL_PMH).setValue(pmh !== null ? pmh : "N/A");
    sheet.getRange(sheetRow, COL_PML).setValue(pml !== null ? pml : "N/A");

    // Evaluate signal
    const arrowCell = sheet.getRange(sheetRow, COL_ARROW);

    if (currentPrice === null || isNaN(pdh) || isNaN(pdl) || isNaN(pmh) || isNaN(pml)) {
      arrowCell.setValue("?").setFontColor("#9E9E9E").setBackground(null);
      return;
    }

    const price = currentPrice;

    const isGreen =
      (price > pmh && price > pdh) ||
      (price >= pdh - NEAR_THRESHOLD) ||
      (price >= pmh - NEAR_THRESHOLD);

    const isRed =
      (price < pml && price < pdl) ||
      (price <= pdl + NEAR_THRESHOLD) ||
      (price <= pml + NEAR_THRESHOLD);

    if (isGreen && !isRed) {
      arrowCell
        .setValue("▲")
        .setFontColor("#00C853")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground("#E8F5E9");
    } else if (isRed && !isGreen) {
      arrowCell
        .setValue("▼")
        .setFontColor("#D50000")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground("#FFEBEE");
    } else {
      arrowCell
        .setValue("—")
        .setFontColor("#FF6F00")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground("#FFF8E1");
    }

    Utilities.sleep(500);
  });

  ss.toast("Actualizado ✓", "Market Tracker", 5);
}

// ── Install time-based trigger ────────────────────────────────
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runNow") ScriptApp.deleteTrigger(t);
  });

  [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY,
   ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY].forEach(day => {
    ScriptApp.newTrigger("runNow")
      .timeBased().atHour(9).nearMinute(31).everyWeeks(1)
      .onWeekDay(day).inTimezone("America/New_York").create();
  });

  SpreadsheetApp.getActiveSpreadsheet()
    .toast("Trigger instalado: Lun-Vie 9:31 ET ✓", "Market Tracker", 8);
}

// ── Menu ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📈 Market Tracker")
    .addItem("1. Setup headers", "setupHeaders")
    .addItem("2. Install trigger (9:31 ET, Lun-Vie)", "installTriggers")
    .addSeparator()
    .addItem("▶ Run", "runNow")
    .addToUi();
}
