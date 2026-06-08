// ============================================================
// MARKET TRACKER - Google Apps Script
// Columns: A=Ticker | B=Signal | C=PDH | D=PDL | E=PMH | F=PML | G=Price
//
// PDH, PDL  → GOOGLEFINANCE (día anterior)
// PMH, PML  → Finnhub candles pre-market 4:00–9:30 AM ET (auto)
// Price     → Finnhub real-time
// ============================================================

const SHEET_NAME     = "Tracker";
const COL_TICKER     = 1; // A
const COL_ARROW      = 2; // B
const COL_PDH        = 3; // C
const COL_PDL        = 4; // D
const COL_PMH        = 5; // E  ← Finnhub auto
const COL_PML        = 6; // F  ← Finnhub auto
const COL_PRICE      = 7; // G  ← Finnhub real-time

// ⚠️ Pegá tu API key de finnhub.io acá:
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
    Logger.log(`[${ticker}] price=${price}`);
    return price > 0 ? price : null;
  } catch (e) {
    Logger.log(`[${ticker}] quote error: ${e}`);
    return null;
  }
}

// ── PMH y PML desde Yahoo Finance (pre-market gratis) ────────
function fetchPreMarketLevels(ticker) {
  // Velas de 1 min del día con pre/post market incluido
  const url = `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(ticker)}?interval=1m&range=1d&includePrePost=true`;
  const headers = { "User-Agent": "Mozilla/5.0" };
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers });
    if (resp.getResponseCode() !== 200) {
      Logger.log(`[${ticker}] Yahoo candles HTTP ${resp.getResponseCode()}`);
      return null;
    }
    const json   = JSON.parse(resp.getContentText());
    const result = json.chart.result[0];
    const quotes = result.indicators.quote[0];
    const timestamps = result.timestamp;

    // Filtrar solo velas del pre-market: antes de las 9:30 AM ET (13:30 UTC en EDT)
    const offsetStr  = Utilities.formatDate(new Date(), "America/New_York", "Z");
    const sign       = offsetStr[0] === "-" ? -1 : 1;
    const offsetMins = sign * (parseInt(offsetStr.slice(1,3)) * 60 + parseInt(offsetStr.slice(3,5)));
    const etDateStr  = Utilities.formatDate(new Date(), "America/New_York", "yyyy-MM-dd");
    const etMidnight = new Date(etDateStr + "T00:00:00Z").getTime() - offsetMins * 60000;
    const preStart   = (etMidnight + 4 * 3600000) / 1000;   // 4:00 AM ET
    const preEnd     = (etMidnight + 9 * 3600000 + 30 * 60000) / 1000; // 9:30 AM ET

    const highs = [], lows = [];
    timestamps.forEach((ts, i) => {
      if (ts >= preStart && ts < preEnd && quotes.high[i] !== null && quotes.low[i] !== null) {
        highs.push(quotes.high[i]);
        lows.push(quotes.low[i]);
      }
    });

    if (highs.length === 0) {
      Logger.log(`[${ticker}] sin velas pre-market en Yahoo`);
      return null;
    }

    const pmh = Math.max(...highs);
    const pml = Math.min(...lows);
    Logger.log(`[${ticker}] PMH=${pmh} PML=${pml} (${highs.length} velas)`);
    return { pmh, pml };
  } catch (e) {
    Logger.log(`[${ticker}] Yahoo candles error: ${e}`);
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

// ── Setup: encabezados + fórmulas ────────────────────────────
function setupHeaders() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const headers = ["Ticker", "Signal", "PDH", "PDL", "PMH", "PML", "Price"];
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
  sheet.setColumnWidth(COL_PMH,    80);
  sheet.setColumnWidth(COL_PML,    80);
  sheet.setColumnWidth(COL_PRICE,  80);

  // Formato numérico 2 decimales para columnas de precios
  const lastRow = Math.max(sheet.getLastRow(), 50);
  [COL_PDH, COL_PDL, COL_PMH, COL_PML, COL_PRICE].forEach(col => {
    sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat("0.00");
  });

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
    [COL_PDH, COL_PDL, COL_PMH, COL_PML, COL_PRICE, COL_ARROW].forEach(col => {
      sheet.getRange(row, col).clearContent().setBackground(null);
    });
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

// ── Run: trae todos los datos y evalúa señales ───────────────
function runNow() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (FINNHUB_API_KEY === "TU_API_KEY_ACÁ") {
    SpreadsheetApp.getUi().alert("⚠️ Pegá tu API key de Finnhub en la línea:\nconst FINNHUB_API_KEY = \"TU_API_KEY_ACÁ\"");
    return;
  }

  if (lastRow < 2) {
    ss.toast("No hay tickers en la columna A.", "Market Tracker", 5);
    return;
  }

  // ── PASADA 1: limpiar señales y asegurar fórmulas solo si la celda está vacía
  for (let row = 2; row <= lastRow; row++) {
    const ticker = sheet.getRange(row, COL_TICKER).getValue().toString().trim();
    if (!ticker) continue;
    // Solo setear fórmula si PDH/PDL están vacíos (evita forzar recálculo)
    if (!sheet.getRange(row, COL_PDH).getValue()) setRowFormulas(sheet, row);
    sheet.getRange(row, COL_ARROW).setValue("⏳").setFontColor("#9E9E9E")
         .setFontSize(14).setHorizontalAlignment("center").setBackground(null);
  }
  SpreadsheetApp.flush();
  Utilities.sleep(2000); // dar tiempo a GOOGLEFINANCE solo si hubo fórmulas nuevas

  // ── PASADA 2: poblar E, F, G y leer C, D para todos los tickers ─
  ss.toast("Paso 1/2 — Cargando PMH, PML, Price…", "Market Tracker", 60);

  const rowData = [];

  for (let i = 0; i <= lastRow - 2; i++) {
    const sheetRow = i + 2;
    const ticker   = sheet.getRange(sheetRow, COL_TICKER).getValue().toString().trim().toUpperCase();
    if (!ticker) continue;

    // PDH/PDL: leer directo (ya deberían tener valor); esperar solo si siguen vacíos
    const pdh = waitForValue(sheet, sheetRow, COL_PDH, 8000);
    const pdl = waitForValue(sheet, sheetRow, COL_PDL, 8000);

    // PMH/PML: Yahoo Finance pre-market candles
    const pm = fetchPreMarketLevels(ticker);
    if (pm) {
      sheet.getRange(sheetRow, COL_PMH).setValue(pm.pmh);
      sheet.getRange(sheetRow, COL_PML).setValue(pm.pml);
    } else {
      sheet.getRange(sheetRow, COL_PMH).setValue("N/A");
      sheet.getRange(sheetRow, COL_PML).setValue("N/A");
    }

    // Price: Finnhub real-time
    const price = fetchFinnhubPrice(ticker);
    if (price !== null) {
      sheet.getRange(sheetRow, COL_PRICE).setValue(price);
    }

    SpreadsheetApp.flush();
    rowData.push({ sheetRow, ticker, pdh, pdl, pmh: pm ? pm.pmh : NaN, pml: pm ? pm.pml : NaN, price });
    Utilities.sleep(300);
  }

  // ── PASADA 3: evaluar señales en columna B ───────────────────
  ss.toast("Paso 2/2 — Evaluando señales…", "Market Tracker", 30);

  let updated = 0;

  rowData.forEach(({ sheetRow, ticker, pdh, pdl, pmh, pml, price }) => {
    const arrowCell = sheet.getRange(sheetRow, COL_ARROW);

    if (pdh === null || pdl === null || price === null) {
      arrowCell.setValue("N/A").setFontColor("#F44336").setFontSize(12)
               .setHorizontalAlignment("center").setBackground(null);
      Logger.log(`${ticker}: sin datos — pdh=${pdh} pdl=${pdl} price=${price}`);
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
