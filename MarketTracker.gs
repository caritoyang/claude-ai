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

// ── 9:30 AM trigger — populate PDH, PDL, PMH, PML ───────────
function run930() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const tickers = sheet.getRange(2, COL_TICKER, lastRow - 1, 1).getValues();

  tickers.forEach((row, i) => {
    const ticker = row[0].toString().trim().toUpperCase();
    if (!ticker) return;

    const data = fetchYahooData(ticker);
    if (!data) return;

    const sheetRow = i + 2;
    sheet.getRange(sheetRow, COL_PDH).setValue(data.pdh !== null ? data.pdh : "N/A");
    sheet.getRange(sheetRow, COL_PDL).setValue(data.pdl !== null ? data.pdl : "N/A");
    sheet.getRange(sheetRow, COL_PMH).setValue(data.pmh !== null ? data.pmh : "N/A");
    sheet.getRange(sheetRow, COL_PML).setValue(data.pml !== null ? data.pml : "N/A");

    // Clear previous arrow
    const arrowCell = sheet.getRange(sheetRow, COL_ARROW);
    arrowCell.setValue("").setBackground(null).setFontColor(null);

    Utilities.sleep(500); // avoid rate limiting
  });

  ss.toast("9:30 data loaded ✓", "Market Tracker", 5);
}

// ── 9:31 AM trigger — evaluate signal arrows ─────────────────
function run931() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, COL_PML).getValues();

  data.forEach((row, i) => {
    const ticker = row[COL_TICKER - 1].toString().trim().toUpperCase();
    if (!ticker) return;

    const pdh = parseFloat(row[COL_PDH - 1]);
    const pdl = parseFloat(row[COL_PDL - 1]);
    const pmh = parseFloat(row[COL_PMH - 1]);
    const pml = parseFloat(row[COL_PML - 1]);

    if (isNaN(pdh) || isNaN(pdl) || isNaN(pmh) || isNaN(pml)) return;

    // Fetch fresh price at 9:31
    const fresh = fetchYahooData(ticker);
    if (!fresh || fresh.currentPrice === null) return;

    const price = fresh.currentPrice;
    const sheetRow = i + 2;
    const arrowCell = sheet.getRange(sheetRow, COL_ARROW);

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
      // Ambiguous — price near both levels
      arrowCell
        .setValue("—")
        .setFontColor("#FF6F00")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground("#FFF8E1");
    }

    Utilities.sleep(500);
  });

  ss.toast("9:31 signals updated ✓", "Market Tracker", 5);
}

// ── Install time-based triggers ──────────────────────────────
function installTriggers() {
  // Remove existing triggers for these functions to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["run930", "run931"].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 9:30 AM ET — runs Mon-Fri, fires between 9:30-9:31
  ScriptApp.newTrigger("run930")
    .timeBased()
    .atHour(9)
    .nearMinute(30)
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .inTimezone("America/New_York")
    .create();

  // Repeat for Tue-Fri
  [ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY,
   ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY].forEach(day => {
    ScriptApp.newTrigger("run930")
      .timeBased().atHour(9).nearMinute(30).everyWeeks(1)
      .onWeekDay(day).inTimezone("America/New_York").create();

    ScriptApp.newTrigger("run931")
      .timeBased().atHour(9).nearMinute(31).everyWeeks(1)
      .onWeekDay(day).inTimezone("America/New_York").create();
  });

  ScriptApp.newTrigger("run931")
    .timeBased().atHour(9).nearMinute(31).everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY).inTimezone("America/New_York").create();

  SpreadsheetApp.getActiveSpreadsheet()
    .toast("Triggers installed for Mon-Fri 9:30 & 9:31 ET ✓", "Market Tracker", 8);
}

// ── Menu ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📈 Market Tracker")
    .addItem("1. Setup headers", "setupHeaders")
    .addItem("2. Install triggers (9:30 & 9:31 ET)", "installTriggers")
    .addSeparator()
    .addItem("▶ Run 9:30 now (manual test)", "run930")
    .addItem("▶ Run 9:31 now (manual test)", "run931")
    .addToUi();
}
