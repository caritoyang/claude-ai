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
  const headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json"
  };

  // Chart endpoint: daily candles with pre/post market, last 5 days
  const chartUrl = `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1d&range=5d&includePrePost=true`;
  let chartResp, chartJson;
  try {
    chartResp = UrlFetchApp.fetch(chartUrl, { muteHttpExceptions: true, headers });
    Logger.log(`[${ticker}] chart HTTP ${chartResp.getResponseCode()}`);
    if (chartResp.getResponseCode() !== 200) {
      Logger.log(`[${ticker}] chart body: ${chartResp.getContentText().substring(0, 300)}`);
      return null;
    }
    chartJson = JSON.parse(chartResp.getContentText());
  } catch (e) {
    Logger.log(`[${ticker}] chart fetch error: ${e}`);
    return null;
  }

  let result;
  try {
    result = chartJson.chart.result[0];
  } catch (e) {
    Logger.log(`[${ticker}] chart parse error: ${e} | body: ${chartResp.getContentText().substring(0, 300)}`);
    return null;
  }

  // Previous completed session high/low
  const timestamps = result.timestamp || [];
  const quotes     = result.indicators.quote[0];
  const today      = new Date();
  today.setHours(0, 0, 0, 0);

  let prevIdx = -1;
  for (let i = timestamps.length - 1; i >= 0; i--) {
    const d = new Date(timestamps[i] * 1000);
    d.setHours(0, 0, 0, 0);
    if (d < today) { prevIdx = i; break; }
  }

  const pdh = prevIdx >= 0 ? quotes.high[prevIdx] : null;
  const pdl = prevIdx >= 0 ? quotes.low[prevIdx]  : null;

  // Current regular market price from chart meta (always present)
  const meta         = result.meta;
  const regularPrice = meta.regularMarketPrice || null;

  // Pre-market data via quoteSummary
  const summaryUrl = `https://query2.finance.yahoo.com/v11/finance/quoteSummary/${encodeURIComponent(ticker)}?modules=price`;
  let pmh = null, pml = null, prePrice = null;
  try {
    const summaryResp = UrlFetchApp.fetch(summaryUrl, { muteHttpExceptions: true, headers });
    Logger.log(`[${ticker}] summary HTTP ${summaryResp.getResponseCode()}`);
    if (summaryResp.getResponseCode() === 200) {
      const summaryJson = JSON.parse(summaryResp.getContentText());
      const pd = summaryJson.quoteSummary.result[0].price;
      pmh      = pd.preMarketHigh  && pd.preMarketHigh.raw  ? pd.preMarketHigh.raw  : null;
      pml      = pd.preMarketLow   && pd.preMarketLow.raw   ? pd.preMarketLow.raw   : null;
      prePrice = pd.preMarketPrice && pd.preMarketPrice.raw ? pd.preMarketPrice.raw : null;
      Logger.log(`[${ticker}] pmh=${pmh} pml=${pml} prePrice=${prePrice} regularPrice=${regularPrice}`);
    }
  } catch (e) {
    Logger.log(`[${ticker}] summary error: ${e}`);
  }

  const currentPrice = prePrice || regularPrice;
  return { pdh, pdl, pmh, pml, currentPrice };
}

// ── Quick test: logs raw data for the first ticker ────────────
function debugFirstTicker() {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const ticker = sheet.getRange(2, COL_TICKER).getValue().toString().trim().toUpperCase();
  if (!ticker) { Logger.log("No ticker in A2"); return; }
  Logger.log(`Testing ticker: ${ticker}`);
  const data = fetchYahooData(ticker);
  Logger.log(`Result: ${JSON.stringify(data)}`);
  SpreadsheetApp.getUi().alert(`${ticker}\n\nPDH: ${data?.pdh}\nPDL: ${data?.pdl}\nPMH: ${data?.pmh}\nPML: ${data?.pml}\nPrecio actual: ${data?.currentPrice}`);
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
    if (!data) {
      sheet.getRange(sheetRow, COL_ARROW).setValue("ERR").setFontColor("#FF0000");
      return;
    }

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

    // Verde: precio mayor a PDH y PMH
    const isGreen = price > pdh && price > pmh;

    // Verde agua: precio por debajo de PDH pero dentro de $5
    const isAqua = !isGreen && price >= pdh - NEAR_THRESHOLD && price <= pdh;

    // Rojo: precio menor a PDL y PML
    const isRed = price < pdl && price < pml;

    // Naranja: precio por encima de PML pero dentro de $5
    const isOrange = !isRed && price <= pml + NEAR_THRESHOLD && price >= pml;

    if (isGreen) {
      arrowCell
        .setValue("▲")
        .setFontColor("#00C853")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground("#E8F5E9");
    } else if (isAqua) {
      arrowCell
        .setValue("▲")
        .setFontColor("#00BCD4")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground("#E0F7FA");
    } else if (isRed) {
      arrowCell
        .setValue("▼")
        .setFontColor("#D50000")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground("#FFEBEE");
    } else if (isOrange) {
      arrowCell
        .setValue("▼")
        .setFontColor("#FF6F00")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground("#FFF3E0");
    } else {
      arrowCell
        .setValue("—")
        .setFontColor("#9E9E9E")
        .setFontSize(16)
        .setHorizontalAlignment("center")
        .setBackground(null);
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
    .addItem("🔍 Debug primer ticker", "debugFirstTicker")
    .addToUi();
}
