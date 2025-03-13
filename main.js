const CHECKBOX_ACTIONS = {
  "Financial Summary": {
    10: {
      2: clearSheet,
      3: fillExampleData,
      4: saveSnapshot,
      5: loadLastSnapshot,
    },
  },
};

const STABLECOIN_CACHE_KEY = "stablecoin_addresses";
const CACHE_DURATION = 21600;

function checkSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Financial Summary");
  if (!sheet) {
    Browser.msgBox("❌ Error", "Financial Summary sheet not found!", Browser.Buttons.OK);
    return;
  }

  let issues = [];

  function checkCellFormat(rowIndex, col, expected) {
    const cell = data[rowIndex][col];
    const format = formats[rowIndex][col];
    const alignment = alignments[rowIndex][col];

    if (!cell) return;

    let text = cell.toString().trim();

    if (expected.isSubItem) {
      if (!text.startsWith("- ")) {
        if (text.startsWith("-")) {
          issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Must have a space after dash "- "`);
        } else {
          issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Sub-items must start with "- "`);
        }
        return;
      }
    }

    if (format.getFontSize() !== expected.fontSize) {
      issues.push(
        `❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Font size must be ${expected.fontSize} (found ${format.getFontSize()})`
      );
    }

    if (expected.isBold && !format.isBold()) {
      issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Text must be bold`);
    }

    if (expected.isItalic && !format.isItalic()) {
      issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Text must be italic`);
    }

    if (alignment !== expected.alignment) {
      issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Must be ${expected.alignment}-aligned`);
    }

    if (text.includes("  ")) {
      issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Contains multiple spaces`);
    }

    if (text !== text.trim()) {
      issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Has leading/trailing spaces`);
    }

    if ((expected.isGroupHeader || expected.isSubItem) && !text.includes("  ")) {
      const textToCheck = expected.isSubItem ? text.substring(2) : text;

      const words = textToCheck.split(/\s+|(?=[()])|(?<=\()/);

      words.forEach((word) => {
        if (word && word !== "(" && word !== ")") {
          if (word[0] !== word[0].toUpperCase()) {
            issues.push(
              `❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Each word must start with capital letter (including "${word}")`
            );
          }
        }
      });
    }

    if (expected.isSubItem && !text.startsWith("- ")) {
      issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Sub-items must start with "- "`);
    }

    if (expected.numberFormat) {
      const numberFormat = sheet.getRange(rowIndex + 1, col + 1).getNumberFormat();

      const allowedFormats = expected.numberFormat === 2 ? ["#,##0.00", "0.00"] : ["#,##0.0000", "0.0000"];

      if (!allowedFormats.includes(numberFormat)) {
        issues.push(
          `❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Cell format must be one of: ${allowedFormats.join(" or ")} (found "${
            numberFormat || "none"
          }")`
        );
      }

      const value = data[rowIndex][col];
      if (value !== "" && value !== null) {
        const numValue = typeof value === "string" ? parseFloat(value.replace(",", ".")) : value;

        if (isNaN(numValue) || typeof numValue !== "number") {
          issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Must be a valid number (found "${value}")`);
        } else {
          const valueString = numValue.toFixed(expected.numberFormat);
          const decimals = valueString.split(".")[1]?.length || 0;
          if (decimals !== expected.numberFormat) {
            sheet.getRange(rowIndex + 1, col + 1).setNumberFormat(allowedFormats[0]);
          }
        }
      }
    }

    if (expected.isCurrency && !/^[A-Z]+$/.test(text)) {
      issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: Currency must be in uppercase letters only`);
    }

    if (expected.isNotes) {
      const words = text.split(" ");
      if (words[0][0] !== words[0][0].toUpperCase()) {
        issues.push(`❌ ${String.fromCharCode(65 + col)}${rowIndex + 1}: First word must be capitalized`);
      }
    }
  }

  const mainHeaders = ["Category", "Amount", "Currency", "Exchange Rate", "To Main Currency", "Notes"];
  const headerRange = sheet.getRange("A1:F1");
  const headerValues = headerRange.getValues()[0];
  const headerFormats = headerRange.getTextStyles();
  const headerAlignments = headerRange.getHorizontalAlignments()[0];

  headerValues.forEach((value, index) => {
    if (value !== mainHeaders[index]) {
      issues.push(`❌ Column ${String.fromCharCode(65 + index)}1: Header mismatch - expected "${mainHeaders[index]}", found "${value}"`);
    }
  });

  headerFormats[0].forEach((format, index) => {
    const col = String.fromCharCode(65 + index);
    if (format.getFontSize() !== 12) {
      issues.push(`❌ Column ${col}1: Font size must be 12 (currently ${format.getFontSize()})`);
    }
    if (!format.isBold()) {
      issues.push(`❌ Column ${col}1: Text must be bold`);
    }
    if (headerAlignments[index] !== "center") {
      issues.push(`❌ Column ${col}1: Must be center-aligned`);
    }
  });

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const formats = dataRange.getTextStyles();
  const alignments = dataRange.getHorizontalAlignments();

  const columnRules = {
    0: {
      groupHeader: { fontSize: 10, isBold: true, alignment: "left", isGroupHeader: true },
      subItem: { fontSize: 10, isBold: false, alignment: "left", isSubItem: true },
      subtotal: { fontSize: 10, isItalic: true, alignment: "left" },
      total: { fontSize: 10, isBold: true, alignment: "left" },
    },
    1: {
      groupHeader: { fontSize: 10, isBold: true, alignment: "right" },
      subItem: { fontSize: 10, alignment: "right", numberFormat: 2 },
      subtotal: { fontSize: 10, isItalic: true, alignment: "right", numberFormat: 2 },
      total: { fontSize: 10, isBold: true, alignment: "right", numberFormat: 2 },
    },
    2: {
      subItem: { fontSize: 10, alignment: "center", isCurrency: true },
      subtotal: { fontSize: 10, isItalic: true, alignment: "center", isCurrency: true },
      total: { fontSize: 10, isBold: true, alignment: "center", isCurrency: true },
    },
    3: {
      subItem: { fontSize: 10, alignment: "center", numberFormat: 4 },
    },
    4: {
      subItem: { fontSize: 10, alignment: "center", numberFormat: 4 },
    },
    5: {
      subItem: { fontSize: 10, isItalic: true, alignment: "left", isNotes: true },
    },
  };

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const cellA = data[rowIndex][0]?.toString().trim() || "";
    const cellAUpper = cellA.toUpperCase();

    const nextRowCellA = data[rowIndex + 1]?.[0]?.toString().trim() || "";
    const isGroupHeader = !cellA.startsWith("- ") && nextRowCellA.startsWith("- ");

    if (cellA === "Subtotal:") {
      for (let col = 0; col < 6; col++) {
        if (columnRules[col]?.subtotal) {
          checkCellFormat(rowIndex, col, columnRules[col].subtotal);
        }
      }
    } else if (cellAUpper === "TOTAL:") {
      for (let col = 0; col < 6; col++) {
        if (columnRules[col]?.total) {
          checkCellFormat(rowIndex, col, columnRules[col].total);
        }
      }
    } else if (cellA.toLowerCase().includes("subtotal") && cellA !== "Subtotal:") {
      issues.push(`❌ ${String.fromCharCode(65)}${rowIndex + 1}: Must be exactly "Subtotal:"`);
    } else if (cellAUpper.includes("TOTAL") && cellAUpper !== "TOTAL:") {
      issues.push(`❌ ${String.fromCharCode(65)}${rowIndex + 1}: Must be exactly "TOTAL:"`);
    } else if (isGroupHeader) {
      for (let col = 0; col < 6; col++) {
        if (columnRules[col]?.groupHeader) {
          checkCellFormat(rowIndex, col, columnRules[col].groupHeader);
        }
      }
    } else {
      const shouldBeSubItem =
        !isGroupHeader && (data[rowIndex - 1]?.[0]?.toString().trim().startsWith("- ") || nextRowCellA.startsWith("- "));

      if (shouldBeSubItem && !cellA.startsWith("- ")) {
        issues.push(`❌ ${String.fromCharCode(65)}${rowIndex + 1}: Sub-items must start with "- "`);
      } else if (cellA.startsWith("- ")) {
        for (let col = 0; col < 6; col++) {
          if (columnRules[col]?.subItem) {
            checkCellFormat(rowIndex, col, columnRules[col].subItem);
          }
        }
      } else if (cellA) {
      }
    }
  }

  const actionHeaders = ["Action", "Description", "Button"];
  const actionHeaderRange = sheet.getRange("H1:J1");
  const actionHeaderValues = actionHeaderRange.getValues()[0];
  const actionHeaderFormats = actionHeaderRange.getTextStyles()[0];
  const actionHeaderAlignments = actionHeaderRange.getHorizontalAlignments()[0];

  actionHeaderValues.forEach((value, index) => {
    const col = String.fromCharCode(72 + index);
    if (value !== actionHeaders[index]) {
      issues.push(`❌ Column ${col}1: Header mismatch - expected "${actionHeaders[index]}", found "${value}"`);
    }
    if (actionHeaderFormats[index].getFontSize() !== 12) {
      issues.push(`❌ Column ${col}1: Font size must be 12`);
    }
    if (!actionHeaderFormats[index].isBold()) {
      issues.push(`❌ Column ${col}1: Text must be bold`);
    }
    if (actionHeaderAlignments[index] !== "center") {
      issues.push(`❌ Column ${col}1: Must be center-aligned`);
    }
  });

  const expectedActions = [
    ["Clear Data", "Remove all data but keep headers", "FALSE"],
    ["Fill Example Data", "Insert sample financial data", "FALSE"],
    ["Save Snapshot", "Save a copy with UTC timestamp", "FALSE"],
    ["Load Last Snapshot", "Restore the last saved snapshot", "FALSE"],
    ["Convert to Main Currency", "Fetch exchange rates and recalculate", "FALSE"],
  ];

  const actionRange = sheet.getRange("H2:J6");
  const actionValues = actionRange.getValues();
  const actionFormats = actionRange.getTextStyles();
  const actionAlignments = actionRange.getHorizontalAlignments();

  actionValues.forEach((row, rowIndex) => {
    const rowNum = rowIndex + 2;

    if (row[0] !== expectedActions[rowIndex][0]) {
      issues.push(`❌ H${rowNum}: Text mismatch - expected "${expectedActions[rowIndex][0]}", found "${row[0]}"`);
    }
    if (actionFormats[rowIndex][0].getFontSize() !== 10) {
      issues.push(`❌ H${rowNum}: Font size must be 10`);
    }
    if (actionAlignments[rowIndex][0] !== "left") {
      issues.push(`❌ H${rowNum}: Must be left-aligned`);
    }
    if (actionFormats[rowIndex][0].isBold() || actionFormats[rowIndex][0].isItalic()) {
      issues.push(`❌ H${rowNum}: Must be normal text (not bold/italic)`);
    }

    if (row[1] !== expectedActions[rowIndex][1]) {
      issues.push(`❌ I${rowNum}: Text mismatch - expected "${expectedActions[rowIndex][1]}", found "${row[1]}"`);
    }
    if (actionFormats[rowIndex][1].getFontSize() !== 10) {
      issues.push(`❌ I${rowNum}: Font size must be 10`);
    }
    if (!actionFormats[rowIndex][1].isItalic()) {
      issues.push(`❌ I${rowNum}: Text must be italic`);
    }
    if (actionAlignments[rowIndex][1] !== "left") {
      issues.push(`❌ I${rowNum}: Must be left-aligned`);
    }

    const buttonValue = row[2].toString().toUpperCase();
    if (buttonValue !== "FALSE") {
      if (buttonValue === "TRUE") {
        issues.push(`⚠️ J${rowNum}: Warning - Checkbox is checked (TRUE)`);
      } else {
        issues.push(`❌ J${rowNum}: Invalid value - expected "FALSE", found "${row[2]}"`);
      }
    }
    if (actionFormats[rowIndex][2].getFontSize() !== 10) {
      issues.push(`❌ J${rowNum}: Font size must be 10`);
    }
    if (actionAlignments[rowIndex][2] !== "center") {
      issues.push(`❌ J${rowNum}: Must be center-aligned`);
    }
  });

  if (issues.length === 0) {
    Browser.msgBox("✅ Structure Check Passed", "All formatting and structure requirements are met!", Browser.Buttons.OK);
  } else {
    const formattedIssues = issues
      .sort((a, b) => {
        const aCell = a.match(/[A-Z]\d+/)[0];
        const bCell = b.match(/[A-Z]\d+/)[0];
        return aCell.localeCompare(bCell);
      })
      .join("\n");

    Browser.msgBox("❌ Structure Check Failed", "The following issues need to be addressed:\n\n" + formattedIssues, Browser.Buttons.OK);
  }
}

function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();

    cache.remove(STABLECOIN_CACHE_KEY);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Financial Summary");
    if (!sheet) throw new Error("Sheet not found");

    const [mainCurrency, totalRowIndex] = findMainCurrency(sheet);
    if (!mainCurrency) throw new Error("Main currency not found");

    const currencies = collectCurrencies(sheet, totalRowIndex);
    const ratesCacheKey = `exchangeRates_${mainCurrency}_${currencies.sort()}`;
    cache.remove(ratesCacheKey);

    Browser.msgBox("Cache cleared successfully!");
  } catch (error) {
    Browser.msgBox("Error clearing cache: " + error.message);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Financial Tools").addItem("Clear Cache", "clearAllCache").addItem("Check Structure", "checkSheetStructure").addToUi();
}

function getStablecoinCache() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(STABLECOIN_CACHE_KEY);
  return cached ? new Set(JSON.parse(cached)) : null;
}

function setStablecoinCache(stablecoins) {
  const cache = CacheService.getScriptCache();
  cache.put(STABLECOIN_CACHE_KEY, JSON.stringify(Array.from(stablecoins)), CACHE_DURATION);
}

function isStablecoin(notes, currency) {
  return notes?.toString().toLowerCase().includes("stablecoin");
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const value = e.value === "TRUE";

  const sheetActions = CHECKBOX_ACTIONS[sheetName]?.[col];

  if (sheetActions && value) {
    const action = sheetActions[row];

    if (action) {
      try {
        e.range.setValue(false);
        SpreadsheetApp.flush();

        action();
      } catch (error) {
        Browser.msgBox(`Error: ${error.message}`);
      }
    }
  }
}

function logToSheet(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Debug Logs");

  if (!logSheet) {
    logSheet = ss.insertSheet("Debug Logs");
    logSheet.getRange("A1:C1").setValues([["Timestamp", "Function", "Message"]]);
    logSheet.setFrozenRows(1);
  }

  const timestamp = new Date().toISOString();
  const caller = new Error().stack.split("\n")[2].trim().split(" ")[1];
  logSheet.appendRow([timestamp, caller, message]);
}

function clearSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Financial Summary");
  if (!sheet) {
    Browser.msgBox("Sheet not found!");
    return;
  }
  var range = sheet.getRange("A2:F");
  range.clearContent();
  Browser.msgBox("Sheet cleared, structure retained!");
}

function fillExampleData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Financial Summary");
  if (!sheet) return;
  clearSheet();

  var data = [
    ["Category", "Amount", "Currency", "Exchange Rate", "To Main Currency", "Notes"],
    ["Bank Accounts", "", "", "", "", ""],
    ["- Bank 1", "1111", "RUB", "87,03", "=B3/D3", "Active funds"],
    ["- Bank 2 (Cashback)", "11", "RUB", "87,03", "=B4/D4", "Pending cashback"],
    ["- Bank 3 (Blocked Funds)", "11", "RUB", "87,03", "=B5/D5", "Harder to withdraw"],
    ["- Bank 4", "111", "EUR", "1,08", "=B6*D6", "Active funds"],
    ["- Bank 4 (Cashback)", "11", "EUR", "1,08", "=B7*D7", "Pending cashback"],
    ["Subtotal:", "=SUM(E3:E7)", "USD", "", "", ""],
    ["Cryptocurrency Holdings", "", "", "", "", ""],
    ["- Crypto 1", "111", "USDT", "1,00", "=B10*D10", "Stablecoin (TON Chain)"],
    ["- Crypto 2", "111", "USDC", "1,00", "=B11*D11", "Stablecoin (BSC Chain)"],
    ["- Crypto 3", "1", "BNB", "300", "=B12*D12", "Binance Coin"],
    ["- Crypto 4", "1", "TON", "3", "=B13*D13", "Toncoin"],
    ["- Crypto 5", "1", "ETH", "2000", "=B14*D14", "Ethereum"],
    ["Subtotal:", "=SUM(E10:E14)", "USD", "", "", ""],
    ["Cash Holdings", "", "", "", "", ""],
    ["- Cash 1", "1111", "EUR", "1,08", "=B17*D17", "Cash on hand"],
    ["Subtotal:", "=SUM(E17)", "USD", "", "", ""],
    ["CS:GO Skins", "", "", "", "", ""],
    ["- Sellable Price", "1111", "USD", "1,00", "=B20*D20", "Steam Market Price: $111"],
    ["Subtotal:", "=SUM(E20)", "USD", "", "", ""],
    ["TOTAL:", "=SUM(E3:E20)", "USD", "", "", ""],
  ];

  var range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);

  var boldRows = [2, 9, 16, 19, 22, 24];
  boldRows.forEach(function (row) {
    sheet.getRange(row, 1, 1, 6).setFontWeight("bold");
  });

  var italicRows = [8, 15, 18, 21, 23];
  italicRows.forEach(function (row) {
    sheet.getRange(row, 1, 1, 6).setFontStyle("italic");
  });

  Browser.msgBox("Example data inserted!");
}

function saveSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Financial Summary");
  if (!sheet) return;

  var now = new Date();
  var timestamp = now.toISOString().replace(/[-:]/g, "").split(".")[0];
  var readableTimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");

  var newSheetName = "Snapshot_" + timestamp;

  var existingSheet = ss.getSheetByName(newSheetName);
  if (existingSheet) ss.deleteSheet(existingSheet);

  var newSheet = sheet.copyTo(ss);
  newSheet.setName(newSheetName);
  ss.setActiveSheet(newSheet);

  var lastRow = newSheet.getLastRow();
  var checkboxRange = newSheet.getRange(1, 10, lastRow, 1);
  checkboxRange.clearContent();
  checkboxRange.setDataValidation(null);

  var headerRange = newSheet.getRange(1, 8, newSheet.getLastRow(), 3);
  headerRange.clearContent();
  headerRange.clearFormat();

  var drawings = newSheet.getDrawings();
  drawings.forEach((drawing) => {
    var pos = drawing.getContainerInfo().getAnchorColumn();
    if (pos >= 8 && pos <= 10) {
      drawing.remove();
    }
  });

  var formattedHeader = [["Date", "Snapshot Creation Time", readableTimestamp]];
  var headerCellRange = newSheet.getRange(1, 8, 1, 3);
  headerCellRange.setValues(formattedHeader).setFontSize(12).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);

  newSheet.getRange(1, 8).setFontWeight("bold");
  newSheet.getRange(1, 9).setFontStyle("italic");
  newSheet.getRange(1, 10).setFontWeight("bold");

  for (var col = 8; col <= 10; col++) {
    newSheet.autoResizeColumn(col);
    newSheet.setColumnWidth(col, newSheet.getColumnWidth(col) + 10);
  }

  Browser.msgBox("Snapshot saved as: " + newSheetName);
}

function loadLastSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss
    .getSheets()
    .map((sheet) => sheet.getName())
    .filter((name) => name.startsWith("Snapshot_"));

  if (sheets.length === 0) {
    Browser.msgBox("No snapshots found!");
    return;
  }

  sheets.sort();
  var lastSnapshotName = sheets[sheets.length - 1];
  var lastSnapshot = ss.getSheetByName(lastSnapshotName);
  var mainSheet = ss.getSheetByName("Financial Summary");

  if (!lastSnapshot || !mainSheet) return;

  clearSheet();

  var lastRow = lastSnapshot.getLastRow();
  var dataRange = lastSnapshot.getRange(1, 1, lastRow, 6);
  var destinationRange = mainSheet.getRange(1, 1, lastRow, 6);

  dataRange.copyTo(destinationRange, { contentsOnly: false });

  Browser.msgBox("Restored from: " + lastSnapshotName);
}

function onError(error) {
  Browser.msgBox(`Critical error: ${error.message}`);
}

function convertToMainCurrency() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Financial Summary");
  if (!sheet) throw new Error("Sheet not found");

  const [mainCurrency, totalRowIndex] = findMainCurrency(sheet);
  if (!mainCurrency) throw new Error("Main currency not found");

  const currencies = collectCurrencies(sheet, totalRowIndex);
  const rates = getExchangeRates(mainCurrency, currencies);

  if (Object.keys(rates).length === 0) {
    throw new Error("Failed to fetch exchange rates");
  }

  updateExchangeData(sheet, mainCurrency, totalRowIndex, rates);
}

function findMainCurrency(sheet) {
  const data = sheet.getDataRange().getValues();
  for (const [i, row] of data.entries()) {
    if (row[0]?.toString().trim().toUpperCase() === "TOTAL:") {
      return [row[2]?.toUpperCase(), i];
    }
  }
  Browser.msgBox("TOTAL: row not found");
  return [null, null];
}

function collectCurrencies(sheet, totalRowIndex) {
  const currencies = new Set();
  let stablecoins = getStablecoinCache();

  if (!stablecoins) {
    stablecoins = new Set();

    const data = sheet.getRange(1, 1, totalRowIndex, 6).getValues();

    for (let i = 1; i < totalRowIndex; i++) {
      const label = data[i][0]?.toString().trim().toUpperCase() || "";
      const currency = data[i][2]?.toString().toUpperCase() || "";
      const notes = data[i][5];

      if (!label.startsWith("SUBTOTAL:") && currency) {
        if (isStablecoin(notes, currency)) {
          stablecoins.add(currency);
        }
        currencies.add(currency);
      }
    }

    setStablecoinCache(stablecoins);
  } else {
    const data = sheet.getRange(1, 1, totalRowIndex, 3).getValues();
    for (let i = 1; i < totalRowIndex; i++) {
      const label = data[i][0]?.toString().trim().toUpperCase() || "";
      const currency = data[i][2]?.toString().toUpperCase() || "";
      if (!label.startsWith("SUBTOTAL:") && currency) {
        currencies.add(currency);
      }
    }
  }

  const finalCurrencies = new Set();
  currencies.forEach((currency) => {
    finalCurrencies.add(stablecoins.has(currency) ? "USD" : currency);
  });

  return Array.from(finalCurrencies);
}

function updateExchangeData(sheet, mainCurrency, totalRowIndex, rates) {
  const dataRange = sheet.getRange(1, 1, totalRowIndex, 6);
  const data = dataRange.getValues();

  let currentSection = [];
  let lastSubtotalRow = -1;

  for (let i = 1; i < totalRowIndex; i++) {
    const rowLabel = data[i][0]?.toString().trim().toUpperCase() || "";

    if (rowLabel.startsWith("SUBTOTAL:")) {
      if (currentSection.length > 0) {
        data[i][1] = `=SUM(E${currentSection[0]}:E${currentSection[currentSection.length - 1]})`;
      }
      if (data[i][2]?.toString().toUpperCase() !== mainCurrency) {
        data[i][2] = mainCurrency;
      }
      lastSubtotalRow = i;
      currentSection = [];
      continue;
    } else if (rowLabel === "TOTAL:") {
      data[i][1] = `=SUM(B2:B${totalRowIndex})`;
      continue;
    }

    if (rowLabel.startsWith("-")) {
      currentSection.push(i + 1);

      const [amount, currency] = parseRowData(data[i]);
      if (amount && currency) {
        const exchangeRate = rates[currency];
        if (exchangeRate) {
          updateRowData(data[i], amount, exchangeRate, mainCurrency, i + 1);
        }
      }
    }
  }

  dataRange.setValues(data);
  Browser.msgBox(`Converted to ${mainCurrency}`);
}

function parseRowData(row) {
  const amount = parseFloat(row[1].toString().replace(",", "."));
  let currency = row[2]?.toUpperCase() || "";
  const notes = row[5];

  const stablecoins = getStablecoinCache();
  if (stablecoins?.has(currency) || isStablecoin(notes, currency)) {
    currency = "USD";
  }

  return [isNaN(amount) ? null : amount, currency || null];
}

function updateRowData(row, amount, rate, baseCurrency, rowIndex) {
  row[3] = rate.toFixed(4).replace(".", ",");
  row[4] = currencyOperationFormula(amount, rate, row[2], baseCurrency, rowIndex);
}

function currencyOperationFormula(amount, rate, currency, baseCurrency, rowNum) {
  const isSameCurrency = currency === baseCurrency;
  return isSameCurrency ? amount : `=B${rowNum}*D${rowNum}`;
}

function getExchangeRates(baseCurrency, currencyList) {
  const cacheKey = `exchangeRates_${baseCurrency}_${currencyList.sort()}`;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  try {
    const rates = {};
    const [fiatRates, cryptoList] = fetchFiatRates(baseCurrency, currencyList);

    Object.assign(rates, fiatRates);
    if (cryptoList.length) {
      Object.assign(rates, fetchCryptoRates(baseCurrency, cryptoList));
    }

    cache.put(cacheKey, JSON.stringify(rates), 3600);
    return rates;
  } catch (e) {
    Browser.msgBox("Error: " + e.message);
    return {};
  }
}

function fetchFiatRates(baseCurrency, currencies) {
  ScriptApp.getOAuthToken();
  const url = `https://open.er-api.com/v6/latest/${baseCurrency}`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());

  if (data.result !== "success") throw new Error("ER-API error");

  const fiatRates = {};
  const cryptoList = [];

  currencies.forEach((curr) => {
    const upperCurr = curr.toUpperCase();
    const rate = data.rates[upperCurr];
    if (rate) {
      fiatRates[upperCurr] = 1 / rate;
    } else {
      cryptoList.push(upperCurr);
    }
  });

  return [fiatRates, cryptoList];
}

function fetchCryptoRates(baseCurrency, cryptoList) {
  ScriptApp.getOAuthToken();
  const assets = JSON.parse(UrlFetchApp.fetch("https://api.coincap.io/v2/assets?limit=2000").getContentText()).data;
  const usdToBaseRate = baseCurrency === "USD" ? 1 : getUsdRate(baseCurrency);
  const rates = assets.reduce((acc, asset) => {
    const symbol = asset.symbol.toUpperCase();
    if (cryptoList.includes(symbol) && asset.priceUsd) {
      acc[symbol] = parseFloat(asset.priceUsd) * usdToBaseRate;
    }
    return acc;
  }, {});
  return rates;
}

function getUsdRate(targetCurrency) {
  ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(`https://open.er-api.com/v6/latest/USD`);
  const data = JSON.parse(response.getContentText());
  if (data.result !== "success") throw new Error("ER-API error");
  const rate = data.rates[targetCurrency.toUpperCase()];
  if (!rate) throw new Error(`USD rate for ${targetCurrency} not found`);
  return rate;
}

function initiateConversion(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName === "Financial Summary" && e.range.getColumn() === 10 && e.range.getRow() === 6 && e.value === "TRUE") {
    try {
      convertToMainCurrency();
      sheet.getRange(6, 11).clearContent();
    } catch (error) {
      Browser.msgBox(`Error: ${error.message}`);
    }

    e.range.setValue(false);
  }
}
