// ============================================================
// MARKET TRACKER - Google Apps Script
// Columns: A=Ticker | B=Signal | C=PDH | D=PDL | E=PMH | F=PML | G=Price
//
// PDH, PDL, Price → fórmulas GOOGLEFINANCE (automáticas)
// PMH, PML        → ingreso manual cada mañana
// ============================================================

const SHEET_NAME     = "Tracker";
const COL_TICKER     = 1; // A
const COL_ARROW      = 2; // B
const COL_PDH        = 3; // C
const COL_PDL        = 4; // D
const COL_PMH        = 5; // E  ← manual
const COL_PML        = 6; // F  ← manual
const COL_PRICE      = 7; // G  (fórmula, puede ocultarse)

const NEAR_THRESHOLD = 5;

// ── Fórmulas GOOGLEFINANCE para una fila ─────────────────────
function setRowFormulas(sheet, row) {
  const a = `A${row}`;
  sheet.getRange(row, COL_PDH)
    .setFormula(`=IFERROR(INDEX(GOOGLEFINANCE(${a},"high",WORKDAY(TODAY(),-1)),2,2),"")`);
  sheet.getRange(row, COL_PDL)
    .setFormula(`=IFERROR(INDEX(GOOGLEFINANCE(${a},"low",WORKDAY(TODAY(),-1)),2,2),"")`);
  sheet.getRange(row, COL_PRICE)
    .setFormula(`=IFERROR(GOOGLEFINANCE(${a},"price"),"")`);
}

// ── Setup: encabezados + fórmulas para todos los tickers ─────
function setupHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  // Encabezados
  const headers = ["Ticker", "Signal", "PDH", "PDL", "PMH ✍", "PML ✍", "Price"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#263238")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Anchos de columna
  sheet.setColumnWidth(COL_TICKER, 90);
  sheet.setColumnWidth(COL_ARROW,  80);
  sheet.setColumnWidth(COL_PDH,    80);
  sheet.setColumnWidth(COL_PDL,    80);
  sheet.setColumnWidth(COL_PMH,    90);
  sheet.setColumnWidth(COL_PML,    90);
  sheet.setColumnWidth(COL_PRICE,  80);

  // Resaltar columnas manuales (PMH/PML) en amarillo claro
  const lastRow = Math.max(sheet.getLastRow(), 10);
  sheet.getRange(2, COL_PMH, lastRow - 1, 1).setBackground("#FFFDE7");
  sheet.getRange(2, COL_PML, lastRow - 1, 1).setBackground("#FFFDE7");

  // Poner fórmulas en las filas que ya tienen tickers
  const lastData = sheet.getLastRow();
  for (let row = 2; row <= lastData; row++) {
    const ticker = sheet.getRange(row, COL_TICKER).getValue().toString().trim();
    if (ticker) setRowFormulas(sheet, row);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Setup completo ✓", "Market Tracker", 4);
}

// ── onEdit: agrega fórmulas automáticamente al escribir ticker
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
    // Si borraron el ticker, limpiar la fila
    sheet.getRange(row, COL_PDH).clearContent();
    sheet.getRange(row, COL_PDL).clearContent();
    sheet.getRange(row, COL_PRICE).clearContent();
    sheet.getRange(row, COL_ARROW).clearContent().setBackground(null);
  }
}

// ── Run: lee valores actuales y evalúa señales ───────────────
function runNow() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ss.toast("No hay tickers en la columna A.", "Market Tracker", 5);
    return;
  }

  // Forzar recálculo antes de leer
  SpreadsheetApp.flush();
  Utilities.sleep(2000); // dar tiempo a GOOGLEFINANCE para resolver

  const data = sheet.getRange(2, 1, lastRow - 1, COL_PRICE).getValues();
  let updated = 0;

  data.forEach((row, i) => {
    const ticker = row[COL_TICKER - 1].toString().trim().toUpperCase();
    if (!ticker) return;

    const pdh   = parseFloat(row[COL_PDH   - 1]);
    const pdl   = parseFloat(row[COL_PDL   - 1]);
    const pmh   = parseFloat(row[COL_PMH   - 1]);
    const pml   = parseFloat(row[COL_PML   - 1]);
    const price = parseFloat(row[COL_PRICE - 1]);

    const sheetRow  = i + 2;
    const arrowCell = sheet.getRange(sheetRow, COL_ARROW);

    // Si faltan datos de GOOGLEFINANCE, esperar
    if (isNaN(pdh) || isNaN(pdl) || isNaN(price)) {
      arrowCell.setValue("⏳").setFontColor("#9E9E9E").setFontSize(14)
               .setHorizontalAlignment("center").setBackground(null);
      return;
    }

    // PMH/PML son opcionales: si no están ingresados, solo se usa PDH/PDL
    const hasPM = !isNaN(pmh) && !isNaN(pml);

    // Verde: precio mayor a PDH (y PMH si está disponible)
    const isGreen = hasPM ? (price > pdh && price > pmh) : (price > pdh);

    // Verde agua: precio dentro de $5 por debajo de PDH
    const isAqua = !isGreen && price >= pdh - NEAR_THRESHOLD && price < pdh;

    // Rojo: precio menor a PDL (y PML si está disponible)
    const isRed = hasPM ? (price < pdl && price < pml) : (price < pdl);

    // Naranja: precio dentro de $5 por encima de PML (solo si PML ingresado)
    const isOrange = hasPM && !isRed && price <= pml + NEAR_THRESHOLD && price >= pml;

    if (isGreen) {
      arrowCell.setValue("▲").setFontColor("#00C853").setFontSize(16)
               .setHorizontalAlignment("center").setBackground("#E8F5E9");
    } else if (isAqua) {
      arrowCell.setValue("▲").setFontColor("#00BCD4").setFontSize(16)
               .setHorizontalAlignment("center").setBackground("#E0F7FA");
    } else if (isRed) {
      arrowCell.setValue("▼").setFontColor("#D50000").setFontSize(16)
               .setHorizontalAlignment("center").setBackground("#FFEBEE");
    } else if (isOrange) {
      arrowCell.setValue("▼").setFontColor("#FF6F00").setFontSize(16)
               .setHorizontalAlignment("center").setBackground("#FFF3E0");
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
