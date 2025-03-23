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

const COLUMN_WIDTHS = {
  1: 220,
  2: 100,
  3: 100,
  4: 150,
  5: 150,
  6: 180,
  7: 20,
  8: 180,
  9: 240,
  10: 80,
};

const ROW_TYPES = {
  MAIN_GROUP: {
    name: "MAIN_GROUP",
    requirements: {
      prefix: "",
      columnRules: {
        A: { required: true, type: "string" },
        B: { required: true, type: "empty" },
        C: { required: true, type: "empty" },
        D: { required: true, type: "empty" },
        E: { required: true, type: "empty" },
        F: { required: true, type: "empty" },
      },
    },
  },
  SUB_GROUP: {
    name: "SUB_GROUP",
    requirements: {
      prefix: "- ",
      columnRules: {
        A: { required: true, type: "string" },
        B: { required: true, type: "number" },
        C: { required: true, type: "string" },
        D: { required: false, type: "number" },
        E: { required: false, type: "number" },
        F: { required: false, type: "string" },
      },
    },
  },
  SUBTOTAL: {
    name: "SUBTOTAL",
    requirements: {
      exactMatch: "Subtotal:",
      columnRules: {
        A: { required: true, type: "string" },
        B: { required: true, type: "number" },
        C: { required: true, type: "string" },
        D: { required: true, type: "empty" },
        E: { required: true, type: "empty" },
        F: { required: true, type: "empty" },
      },
    },
  },
  TOTAL: {
    name: "TOTAL",
    requirements: {
      exactMatch: "TOTAL:",
      columnRules: {
        A: { required: true, type: "string" },
        B: { required: true, type: "number" },
        C: { required: true, type: "string" },
        D: { required: true, type: "empty" },
        E: { required: true, type: "empty" },
        F: { required: true, type: "empty" },
      },
    },
  },
};

function checkInternalStructure(sheet) {
  const errors = [];

  try {
    const totalRow = findTotalRow(sheet);
    if (!totalRow) {
      errors.push('Missing "TOTAL:" row');
      return [false, errors];
    }

    const structure = parseSheetStructure(sheet, 2, totalRow);

    validateStructure(structure, errors);

    return [errors.length === 0, errors];
  } catch (error) {
    errors.push(`Critical error: ${error.message}`);
    return [false, errors];
  }
}

function findTotalRow(sheet) {
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const values = range.getValues();
  let found = false;
  let firstTotalRow = null;

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "TOTAL:") {
      if (!found) {
        found = true;
        firstTotalRow = i + 2;
      } else {
        throw new Error("Multiple TOTAL: rows found. Only one is allowed.");
      }
    }
  }
  return firstTotalRow;
}

function isCellEmpty(cell) {
  const value = cell.getValue();
  const formula = cell.getFormula();
  if (value !== "" || formula !== "") {
    return "Cell contains value or formula";
  }

  const textStyle = cell.getTextStyle();
  if (textStyle.isBold()) return "Cell has bold formatting";
  if (textStyle.isItalic()) return "Cell has italic formatting";
  if (textStyle.isUnderline()) return "Cell has underline formatting";
  if (textStyle.isStrikethrough()) return "Cell has strikethrough formatting";

  if (cell.getFontFamily() !== "Arial") {
    return `Wrong font family: ${cell.getFontFamily()} (should be Arial)`;
  }
  if (cell.getFontSize() !== 10) {
    return `Wrong font size: ${cell.getFontSize()} (should be 10)`;
  }
  if (cell.getFontColor() !== "#000000") {
    return `Wrong font color: ${cell.getFontColor()} (should be #000000)`;
  }

  if (cell.getHorizontalAlignment() !== "center") {
    return "Cell is not center-aligned horizontally";
  }
  if (cell.getVerticalAlignment() !== "middle") {
    return "Cell is not center-aligned vertically";
  }

  if (cell.getBackground() !== "#ffffff") {
    return `Wrong background color: ${cell.getBackground()} (should be #ffffff)`;
  }

  const richText = cell.getRichTextValue();
  if (richText && richText.getText() !== "") {
    return "Cell contains rich text formatting";
  }

  const note = cell.getNote();
  if (note && note !== "") {
    return "Cell contains note";
  }

  const dataValidation = cell.getDataValidation();
  if (dataValidation !== null) {
    return "Cell contains data validation";
  }

  if (formula.toLowerCase().includes("=hyperlink(")) {
    return "Cell contains hyperlink";
  }

  const sheet = cell.getSheet();
  const drawings = sheet.getDrawings();
  const cellColumn = cell.getColumn();
  const cellRow = cell.getRow();

  for (const drawing of drawings) {
    const anchor = drawing.getContainerInfo();
    if (anchor.getAnchorColumn() === cellColumn - 1 && anchor.getAnchorRow() === cellRow - 1) {
      return "Cell contains drawing or image";
    }
  }

  const rules = sheet.getConditionalFormatRules();
  for (const rule of rules) {
    const ranges = rule.getRanges();
    for (const range of ranges) {
      if (range.getA1Notation().includes(cell.getA1Notation())) {
        return "Cell has conditional formatting";
      }
    }
  }

  return true;
}

function parseSheetStructure(sheet, startRow, totalRow) {
  const structure = {
    groups: [],
    totalRow: totalRow,
  };

  let currentRow = startRow;
  while (currentRow < totalRow) {
    const rowValues = sheet.getRange(currentRow, 1, 1, 6).getValues()[0];
    const category = rowValues[0].toString();

    if (!category.startsWith("-") && category !== "Subtotal:") {
      const group = parseGroup(sheet, currentRow, totalRow);
      structure.groups.push(group);
      currentRow = group.endRow + 1;
    } else {
      currentRow++;
    }
  }

  return structure;
}

function parseGroup(sheet, startRow, totalRow) {
  const group = {
    name: sheet.getRange(startRow, 1).getValue(),
    subgroups: [],
    startRow: startRow,
    endRow: null,
  };

  let currentRow = startRow + 1;
  while (currentRow < totalRow) {
    const value = sheet.getRange(currentRow, 1).getValue().toString();

    if (value === "Subtotal:") {
      group.endRow = currentRow;
      break;
    }

    if (value.startsWith("-")) {
      group.subgroups.push({
        name: value,
        row: currentRow,
      });
    } else if (value && !value.startsWith("-")) {
      currentRow--;
      group.endRow = currentRow;
      break;
    }

    currentRow++;
  }

  return group;
}

function validateTextFormatting(structure, errors) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Financial Summary");

  const totalRow = structure.totalRow;
  validateTotalRowFormatting(sheet, totalRow, errors);

  structure.groups.forEach((group) => {
    validateMainGroupFormatting(sheet, group.startRow, errors);

    group.subgroups.forEach((subgroup) => {
      validateSubgroupFormatting(sheet, subgroup.row, errors);
    });

    if (group.endRow) {
      validateSubtotalFormatting(sheet, group.endRow, errors);
    }
  });
}

function validateTotalRowFormatting(sheet, row, errors) {
  validateCellFormatting(
    sheet,
    row,
    1,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: true,
      isItalic: false,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "left",
      verticalAlignment: "middle",
    },
    errors
  );

  validateCellFormatting(
    sheet,
    row,
    2,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: true,
      isItalic: false,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "right",
      verticalAlignment: "middle",
    },
    errors
  );

  validateCellFormatting(
    sheet,
    row,
    3,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: true,
      isItalic: false,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "center",
      verticalAlignment: "middle",
    },
    errors
  );
}

function validateMainGroupFormatting(sheet, row, errors) {
  validateCellFormatting(
    sheet,
    row,
    1,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: true,
      isItalic: false,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "left",
      verticalAlignment: "middle",
    },
    errors
  );
}

function validateSubgroupFormatting(sheet, row, errors) {
  validateCellFormatting(
    sheet,
    row,
    1,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: false,
      isItalic: false,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "left",
      verticalAlignment: "middle",
    },
    errors
  );

  validateCellFormatting(
    sheet,
    row,
    2,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: false,
      isItalic: false,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "right",
      verticalAlignment: "middle",
    },
    errors
  );

  validateCellFormatting(
    sheet,
    row,
    3,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: false,
      isItalic: false,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "center",
      verticalAlignment: "middle",
    },
    errors
  );

  [4, 5].forEach((col) => {
    const cell = sheet.getRange(row, col);
    if (cell.getValue()) {
      validateCellFormatting(
        sheet,
        row,
        col,
        {
          fontFamily: "Arial",
          fontSize: 10,
          isBold: false,
          isItalic: false,
          isUnderline: false,
          isStrikethrough: false,
          horizontalAlignment: "left",
          verticalAlignment: "middle",
        },
        errors
      );
    }
  });

  const cellF = sheet.getRange(row, 6);
  if (cellF.getValue()) {
    validateCellFormatting(
      sheet,
      row,
      6,
      {
        fontFamily: "Arial",
        fontSize: 10,
        isBold: false,
        isItalic: true,
        isUnderline: false,
        isStrikethrough: false,
        horizontalAlignment: "left",
        verticalAlignment: "middle",
      },
      errors
    );
  }
}

function validateSubtotalFormatting(sheet, row, errors) {
  validateCellFormatting(
    sheet,
    row,
    1,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: false,
      isItalic: true,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "left",
      verticalAlignment: "middle",
    },
    errors
  );

  validateCellFormatting(
    sheet,
    row,
    2,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: false,
      isItalic: true,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "right",
      verticalAlignment: "middle",
    },
    errors
  );

  validateCellFormatting(
    sheet,
    row,
    3,
    {
      fontFamily: "Arial",
      fontSize: 10,
      isBold: false,
      isItalic: true,
      isUnderline: false,
      isStrikethrough: false,
      horizontalAlignment: "center",
      verticalAlignment: "middle",
    },
    errors
  );
}

function validateCellFormatting(sheet, row, col, expected, errors) {
  const cell = sheet.getRange(row, col);
  const textStyle = cell.getTextStyle();
  const cellAddress = `${columnToLetter(col)}${row}`;

  if (cell.getFontFamily() !== expected.fontFamily) {
    errors.push(`${cellAddress}: Wrong font family (should be ${expected.fontFamily})`);
  }

  if (cell.getFontSize() !== expected.fontSize) {
    errors.push(`${cellAddress}: Wrong font size (should be ${expected.fontSize})`);
  }

  if (textStyle.isBold() !== expected.isBold) {
    errors.push(`${cellAddress}: ${expected.isBold ? "Should be bold" : "Should not be bold"}`);
  }

  if (textStyle.isItalic() !== expected.isItalic) {
    errors.push(`${cellAddress}: ${expected.isItalic ? "Should be italic" : "Should not be italic"}`);
  }

  if ((cell.getFontLine() === "underline") !== expected.isUnderline) {
    errors.push(`${cellAddress}: ${expected.isUnderline ? "Should be underlined" : "Should not be underlined"}`);
  }

  if ((cell.getFontLine() === "line-through") !== expected.isStrikethrough) {
    errors.push(`${cellAddress}: ${expected.isStrikethrough ? "Should be strikethrough" : "Should not be strikethrough"}`);
  }

  if (cell.getHorizontalAlignment().toLowerCase() !== expected.horizontalAlignment) {
    errors.push(`${cellAddress}: Wrong horizontal alignment (should be ${expected.horizontalAlignment})`);
  }

  if (cell.getVerticalAlignment().toLowerCase() !== expected.verticalAlignment) {
    errors.push(`${cellAddress}: Wrong vertical alignment (should be ${expected.verticalAlignment})`);
  }
}

function validateEmptyRowsAfterTotal(structure, errors) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Financial Summary");
  const startRow = structure.totalRow + 1;
  const endRow = 1000;
  const range = sheet.getRange(startRow, 1, endRow - startRow + 1, 6);

  const values = range.getValues();
  const textStyles = range.getTextStyles();
  const backgrounds = range.getBackgrounds();
  const horizontalAlignments = range.getHorizontalAlignments();
  const verticalAlignments = range.getVerticalAlignments();
  const fontFamilies = range.getFontFamilies();
  const fontSizes = range.getFontSizes();
  const fontColors = range.getFontColors();
  const dataValidations = range.getDataValidations();
  const notes = range.getNotes();
  const formulas = range.getFormulas();

  const rules = sheet.getConditionalFormatRules();
  const rulesAffectingRange = rules.filter((rule) => {
    return rule.getRanges().some((r) => {
      const a1 = r.getA1Notation();
      return a1.includes(`${startRow}:`) || a1.includes(`:${endRow}`);
    });
  });

  const drawings = sheet.getDrawings();
  const drawingsInRange = drawings.filter((drawing) => {
    const anchor = drawing.getContainerInfo();
    const row = anchor.getAnchorRow() + 1;
    return row >= startRow && row <= endRow;
  });

  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < 6; j++) {
      const actualRow = startRow + i;
      const columnLetter = columnToLetter(j + 1);
      const cellAddress = `${columnLetter}${actualRow}`;

      if (values[i][j] !== "" || formulas[i][j] !== "") {
        errors.push(`${cellAddress}: Cell contains value or formula`);
        continue;
      }

      const textStyle = textStyles[i][j];
      if (textStyle.isBold()) {
        errors.push(`${cellAddress}: Cell has bold formatting`);
      }
      if (textStyle.isItalic()) {
        errors.push(`${cellAddress}: Cell has italic formatting`);
      }
      if (textStyle.isUnderline()) {
        errors.push(`${cellAddress}: Cell has underline formatting`);
      }
      if (textStyle.isStrikethrough()) {
        errors.push(`${cellAddress}: Cell has strikethrough formatting`);
      }

      if (fontFamilies[i][j] !== "Arial") {
        errors.push(`${cellAddress}: Wrong font family (should be Arial)`);
      }
      if (fontSizes[i][j] !== 10) {
        errors.push(`${cellAddress}: Wrong font size (should be 10)`);
      }
      if (fontColors[i][j] !== "#000000") {
        errors.push(`${cellAddress}: Wrong font color (should be #000000)`);
      }

      if (horizontalAlignments[i][j] !== "center") {
        errors.push(`${cellAddress}: Cell is not center-aligned horizontally`);
      }
      if (verticalAlignments[i][j] !== "middle") {
        errors.push(`${cellAddress}: Cell is not center-aligned vertically`);
      }

      if (backgrounds[i][j] !== "#ffffff") {
        errors.push(`${cellAddress}: Wrong background color (should be #ffffff)`);
      }

      if (notes[i][j] !== "") {
        errors.push(`${cellAddress}: Cell contains note`);
      }

      if (dataValidations[i][j] !== null) {
        errors.push(`${cellAddress}: Cell contains data validation`);
      }

      if (rulesAffectingRange.length > 0) {
        errors.push(`${cellAddress}: Cell has conditional formatting`);
      }

      if (
        drawingsInRange.some((drawing) => {
          const anchor = drawing.getContainerInfo();
          return anchor.getAnchorRow() + 1 === actualRow && anchor.getAnchorColumn() + 1 === j + 1;
        })
      ) {
        errors.push(`${cellAddress}: Cell contains drawing or image`);
      }
    }
  }
}

function validateStructure(structure, errors) {
  validateRow(structure.totalRow, ROW_TYPES.TOTAL, errors);

  structure.groups.forEach((group, index) => {
    validateGroup(group, index === 0, errors);
  });

  validateEmptyRowsAfterTotal(structure, errors);

  validateTextFormatting(structure, errors);
}

function validateGroup(group, isFirst, errors) {
  validateRow(group.startRow, ROW_TYPES.MAIN_GROUP, errors);

  let lastSubgroupRow = group.startRow;
  group.subgroups.forEach((subgroup) => {
    if (subgroup.row !== lastSubgroupRow + 1) {
      errors.push(`Empty line detected between subgroups at row ${lastSubgroupRow + 1}`);
    }
    validateRow(subgroup.row, ROW_TYPES.SUB_GROUP, errors);
    lastSubgroupRow = subgroup.row;
  });

  if (group.endRow) {
    validateRow(group.endRow, ROW_TYPES.SUBTOTAL, errors);
  } else {
    errors.push(`Missing Subtotal for group "${group.name}"`);
  }
}

function validateRow(row, rowType, errors) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Financial Summary");
  const range = sheet.getRange(row, 1, 1, 6);
  const values = range.getValues()[0];
  const numberFormats = range.getNumberFormats()[0];

  const rules = rowType.requirements;

  if (rules.exactMatch && values[0] !== rules.exactMatch) {
    errors.push(`Row ${row}: Expected "${rules.exactMatch}", got "${values[0]}"`);
    return;
  }

  if (rules.prefix && !values[0].startsWith(rules.prefix)) {
    errors.push(`Row ${row}: Should start with "${rules.prefix}"`);
    return;
  }

  Object.entries(rules.columnRules).forEach(([col, rule]) => {
    const colIndex = col.charCodeAt(0) - 65;
    const cell = range.getCell(1, colIndex + 1);
    const format = numberFormats[colIndex];

    if (rule.type === "empty") {
      const emptinessCheck = isCellEmpty(cell);
      if (emptinessCheck !== true) {
        errors.push(`${col}${row}: ${emptinessCheck}`);
      }
      return;
    }

    const value = values[colIndex];

    if (rule.empty && value !== "") {
      errors.push(`${col}${row}: Should be empty`);
      return;
    }

    if (rule.required && !value) {
      errors.push(`${col}${row}: Required value missing`);
      return;
    }

    if (value && rule.type) {
      if (rule.type === "number") {
        if (typeof value !== "number") {
          errors.push(`${col}${row}: Should be a number`);
        }
        const validFormats = ["0.0000", "0,0000", "#,####", "#.####"];
        if (!validFormats.includes(format)) {
          errors.push(`${col}${row}: Wrong number format (should be one of: ${validFormats.join(", ")})`);
        }
      } else if (rule.type === "string" && typeof value !== "string") {
        errors.push(`${col}${row}: Should be text`);
      }
    }
  });
}

function checkSheetStructure() {
  const errors = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const [initialOk, initialErrors] = checkInitialStructure(ss);
  if (!initialOk) errors.push(...initialErrors);

  const financialSummary = ss.getSheetByName("Financial Summary");
  if (financialSummary) {
    const [externalOk, externalErrors] = checkExternalStructure(financialSummary);
    if (!externalOk) errors.push(...externalErrors);

    const [internalOk, internalErrors] = checkInternalStructure(financialSummary);
    if (!internalOk) errors.push(...internalErrors);
  }

  if (errors.length === 0) {
    Browser.msgBox("✅ Structure Check", "All structure requirements are met!", Browser.Buttons.OK);
  } else {
    Browser.msgBox("❌ Structure Check Failed", errors.join("\n"), Browser.Buttons.OK);
  }

  return {
    ok: errors.length === 0,
    errors: errors.length > 0 ? errors : null,
  };
}

function restoreInternalStructure(sheet) {
  const totalRow = findTotalRow(sheet);
  if (!totalRow) return;

  const structure = parseSheetStructure(sheet, 2, totalRow);

  const dataRange = sheet.getRange(2, 1, totalRow - 1, 6);
  dataRange
    .setFontFamily("Arial")
    .setFontSize(10)
    .setFontColor("#000000")
    .setBackground("#ffffff")
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  const totalRowRange = sheet.getRange(totalRow, 1, 1, 3);
  totalRowRange.setFontWeight("bold");
  sheet.getRange(totalRow, 1).setHorizontalAlignment("left");
  sheet.getRange(totalRow, 2).setHorizontalAlignment("right");
  sheet.getRange(totalRow, 3).setHorizontalAlignment("center");

  sheet.getRange(totalRow, 2).setNumberFormat("0.0000");

  structure.groups.forEach((group) => {
    const mainGroupRange = sheet.getRange(group.startRow, 1);
    mainGroupRange.setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");

    group.subgroups.forEach((subgroup) => {
      const subgroupRow = sheet.getRange(subgroup.row, 1, 1, 6);

      subgroupRow.getCell(1, 1).setFontWeight("normal").setHorizontalAlignment("left").setVerticalAlignment("middle");

      subgroupRow
        .getCell(1, 2)
        .setFontWeight("normal")
        .setHorizontalAlignment("right")
        .setVerticalAlignment("middle")
        .setNumberFormat("0.0000");

      subgroupRow.getCell(1, 4).setNumberFormat("0.0000");

      subgroupRow.getCell(1, 5).setNumberFormat("0.0000");

      subgroupRow.getCell(1, 3).setFontWeight("normal").setHorizontalAlignment("center").setVerticalAlignment("middle");

      subgroupRow.getCell(1, 4).setHorizontalAlignment("left").setVerticalAlignment("middle");

      subgroupRow.getCell(1, 5).setHorizontalAlignment("left").setVerticalAlignment("middle");

      const notesCell = subgroupRow.getCell(1, 6);
      if (notesCell.getValue()) {
        notesCell.setFontStyle("italic").setHorizontalAlignment("left").setVerticalAlignment("middle");
      }
    });

    if (group.endRow) {
      const subtotalRange = sheet.getRange(group.endRow, 1, 1, 3);
      subtotalRange.setFontStyle("italic").setVerticalAlignment("middle");

      sheet.getRange(group.endRow, 1).setHorizontalAlignment("left");

      sheet.getRange(group.endRow, 2).setHorizontalAlignment("right").setNumberFormat("0.0000");

      sheet.getRange(group.endRow, 3).setHorizontalAlignment("center");
    }
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > totalRow) {
    const clearRange = sheet.getRange(totalRow + 1, 1, lastRow - totalRow, 6);
    clearRange
      .clear()
      .setFontFamily("Arial")
      .setFontSize(10)
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontColor("#000000")
      .setBackground("#ffffff")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  }
}

function checkInitialStructure(spreadsheet) {
  const errors = [];
  const sheets = spreadsheet.getSheets();
  const sheetNames = sheets.map((sheet) => sheet.getName());

  const hasFinancialSummary = sheetNames.includes("Financial Summary");
  const snapshotSheets = sheetNames.filter((name) => name.startsWith("Snapshot_") && name !== "Financial Summary");
  const hasDebugLogs = sheetNames.includes("Debug Logs");

  if (!hasFinancialSummary) {
    errors.push('Missing required sheet: "Financial Summary"');
  }

  const allowedSheetsCount = (hasFinancialSummary ? 1 : 0) + snapshotSheets.length + (hasDebugLogs ? 1 : 0);

  if (sheetNames.length !== allowedSheetsCount) {
    errors.push('Invalid sheets detected. Only "Financial Summary", "Debug Logs" and Snapshot_ sheets allowed');
  }

  return [errors.length === 0, errors];
}

function restoreInitialStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const messages = [];

  const financialSummary = ss.getSheetByName("Financial Summary");
  if (!financialSummary) {
    ss.insertSheet("Financial Summary");
    messages.push('Created "Financial Summary" sheet');
  }

  sheets.forEach((sheet) => {
    const name = sheet.getName();
    if (name !== "Financial Summary" && name !== "Debug Logs" && !name.startsWith("Snapshot_")) {
      ss.deleteSheet(sheet);
      messages.push(`Removed invalid sheet: "${name}"`);
    }
  });

  if (messages.length === 0) {
    Browser.msgBox("✅ Structure Check", "Financial Summary sheet is already in place!", Browser.Buttons.OK);
  } else {
    Browser.msgBox("✅ Structure Restored", messages.join("\n"), Browser.Buttons.OK);
  }
}

function restoreExternalStructure(sheet) {
  const mainHeaders = [
    ["Category", "Amount", "Currency", "Exchange Rate", "To Main Currency", "Notes", "", "Action", "Description", "Button"],
  ];
  const headerRange = sheet.getRange(1, 1, 1, 10);
  headerRange.setValues(mainHeaders);

  const rangeA1F1000 = sheet.getRange("A1:F1000");
  const rangeH1J6 = sheet.getRange("H1:J6");
  const rangeG1G1000 = sheet.getRange("G1:G1000");
  const rangeF1F1000 = sheet.getRange("F1:F1000");
  const rangeH1H6 = sheet.getRange("H1:H6");

  rangeA1F1000
    .setFontFamily("Arial")
    .setFontColor("#000000")
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID)
    .setBackground(null);

  rangeH1J6
    .setFontFamily("Arial")
    .setFontColor("#000000")
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID)
    .setBackground(null);

  rangeG1G1000.setBorder(false, false, false, false, false, false).setBackground(null).setFontFamily("Arial");

  rangeF1F1000.setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);

  rangeH1H6.setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);

  sheet
    .getRange("A1:F1")
    .setFontFamily("Arial")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet
    .getRange("H1:J1")
    .setFontFamily("Arial")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  const actionData = [
    ["Clear Data", "Remove all data but keep headers", ""],
    ["Fill Example Data", "Insert sample financial data", ""],
    ["Save Snapshot", "Save a copy with UTC timestamp", ""],
    ["Load Last Snapshot", "Restore the last saved snapshot", ""],
    ["Convert to Main Currency", "Fetch exchange rates and recalculate", ""],
  ];

  const actionRange = sheet.getRange(2, 8, 5, 3);
  actionRange.setValues(actionData);

  const actionLabels = sheet.getRange(2, 8, 5, 1);
  actionLabels
    .setFontFamily("Arial")
    .setFontSize(10)
    .setFontWeight("normal")
    .setFontStyle("normal")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  const actionDescriptions = sheet.getRange(2, 9, 5, 1);
  actionDescriptions
    .setFontFamily("Arial")
    .setFontSize(10)
    .setFontWeight("normal")
    .setFontStyle("italic")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  const checkboxRange = sheet.getRange(2, 10, 5, 1);
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  checkboxRange
    .setDataValidation(rule)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontSize(10)
    .setFontWeight("normal")
    .setFontStyle("normal");

  Object.entries(COLUMN_WIDTHS).forEach(([col, width]) => {
    sheet.setColumnWidth(parseInt(col), width);
  });
}

function checkExternalStructure(sheet) {
  const errors = [];
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  const expectedHeaders = [
    "Category",
    "Amount",
    "Currency",
    "Exchange Rate",
    "To Main Currency",
    "Notes",
    "",
    "Action",
    "Description",
    "Button",
  ];

  expectedHeaders.forEach((expected, col) => {
    if (expected && values[0][col] !== expected) {
      errors.push(`${getCellAddress(1, col + 1)}: Should contain "${expected}"`);
    }
  });

  const mainHeaders = sheet.getRange("A1:F1");
  const actionHeaders = sheet.getRange("H1:J1");

  checkFormatting(sheet, errors);

  checkHeaderFormatting(mainHeaders, errors);
  checkHeaderFormatting(actionHeaders, errors);

  checkButtons(sheet, errors);

  return [errors.length === 0, errors];
}

function checkHeaderFormatting(range, errors) {
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numCols = range.getNumColumns();

  for (let col = 0; col < numCols; col++) {
    const cell = range.getCell(1, col + 1);
    const address = getCellAddress(startRow, startCol + col);
    const size = cell.getFontSize();
    if (size !== 12) {
      errors.push(`${address}: Wrong font size (should be 12)`);
    }

    const isBold = cell.getFontWeight() === "bold";
    if (!isBold) {
      errors.push(`${address}: Should be bold`);
    }

    const unwantedStyles = [];

    if (cell.getFontStyle() !== "normal") {
      unwantedStyles.push("italic");
    }

    if (cell.getFontLine() === "line-through") {
      unwantedStyles.push("strikethrough");
    }

    if (cell.getFontLine() === "underline") {
      unwantedStyles.push("underline");
    }

    if (unwantedStyles.length > 0) {
      errors.push(`${address}: Found unwanted styles (${unwantedStyles.join(", ")})`);
    }

    const hAlign = cell.getHorizontalAlignment().toLowerCase();
    const vAlign = cell.getVerticalAlignment().toLowerCase();

    if (hAlign !== "center") {
      errors.push(`${address}: Should be center-aligned horizontally`);
    }

    if (vAlign !== "middle") {
      errors.push(`${address}: Should be center-aligned vertically`);
    }
  }
}

function checkFormatting(sheet, errors) {
  const rangeA1F1000 = sheet.getRange("A1:F1000");
  const rangeH1J6 = sheet.getRange("H1:J6");
  const rangeG1G1000 = sheet.getRange("G1:G1000");

  checkNoBackground(rangeA1F1000, errors);
  checkNoBackground(rangeH1J6, errors);

  checkFontAndColor(rangeA1F1000, errors);
  checkFontAndColor(rangeH1J6, errors);

  checkBorders(rangeA1F1000, errors);
  checkBorders(rangeH1J6, errors);
  checkGColumnContent(rangeG1G1000, errors);
  checkColumnWidths(sheet, errors);
}

function checkColumnWidths(sheet, errors) {
  Object.entries(COLUMN_WIDTHS).forEach(([col, expectedWidth]) => {
    const columnLetter = columnToLetter(parseInt(col));
    const actualWidth = sheet.getColumnWidth(parseInt(col));

    const minWidth = expectedWidth - 5;
    const maxWidth = expectedWidth + 5;

    if (actualWidth < minWidth || actualWidth > maxWidth) {
      errors.push(`Column ${columnLetter}: Wrong width (current: ${actualWidth}px, expected: ${expectedWidth}px)`);
    }
  });
}

function checkGColumnContent(range, errors) {
  const values = range.getValues();
  const validations = range.getDataValidations();
  const notes = range.getNotes();
  const fontFamilies = range.getFontFamilies();
  const fontColors = range.getFontColors();
  const backgrounds = range.getBackgrounds();
  const textStyles = range.getTextStyles();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = range.getSheet();
  const fileId = ss.getId();
  const sheetName = sheet.getName();
  const token = ScriptApp.getOAuthToken();
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${fileId}?ranges=${encodeURIComponent(
    sheetName + "!G1:G1000"
  )}&fields=sheets/data/rowData/values/userEnteredFormat/borders`;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true,
    });
    const result = JSON.parse(response.getContentText());
    const rowData = result.sheets?.[0]?.data?.[0]?.rowData || [];

    values.forEach((row, rowIndex) => {
      const actualRow = rowIndex + 1;
      const cellAddress = `G${actualRow}`;

      if (row[0] !== "") {
        errors.push(`${cellAddress}: Should be empty`);
      }

      if (validations[rowIndex][0] !== null) {
        errors.push(`${cellAddress}: Should not have data validation`);
      }

      if (notes[rowIndex][0] !== "") {
        errors.push(`${cellAddress}: Should not have notes`);
      }

      if (fontFamilies[rowIndex][0] !== "Arial") {
        errors.push(`${cellAddress}: Should have default font (Arial)`);
      }
      if (fontColors[rowIndex][0] !== "#000000") {
        errors.push(`${cellAddress}: Should have default text color (black)`);
      }
      if (backgrounds[rowIndex][0] !== "#ffffff") {
        errors.push(`${cellAddress}: Should have no background color`);
      }

      const textStyle = textStyles[rowIndex][0];
      if (textStyle.isBold() || textStyle.isItalic() || textStyle.isUnderline() || textStyle.isStrikethrough()) {
        errors.push(`${cellAddress}: Should not have any text styling`);
      }

      const cellBorders = rowData[rowIndex]?.values?.[0]?.userEnteredFormat?.borders;
      if (cellBorders) {
        if (cellBorders.top?.style !== "NONE") {
          errors.push(`${cellAddress}: Should not have top border`);
        }
        if (cellBorders.bottom?.style !== "NONE") {
          errors.push(`${cellAddress}: Should not have bottom border`);
        }
      }
    });

    const drawings = sheet.getDrawings();
    drawings.forEach((drawing) => {
      const anchorCol = drawing.getContainerInfo().getAnchorColumn();
      if (anchorCol === 6) {
        errors.push(`Column G: Should not contain any drawings or objects`);
      }
    });
  } catch (error) {
    errors.push(`Column G: Critical error checking format - ${error.message}`);
  }
}

function checkNoBackground(range, errors) {
  const backgrounds = range.getBackgrounds();
  backgrounds.forEach((row, rIndex) => {
    row.forEach((bg, cIndex) => {
      if (bg !== "#ffffff") {
        const cellAddress = getCellAddress(rIndex + range.getRow(), cIndex + range.getColumn());
        errors.push(`${cellAddress}: Should have no background color`);
      }
    });
  });
}

function checkFontAndColor(range, errors) {
  const fonts = range.getFontFamilies();
  const colors = range.getFontColors();

  fonts.forEach((row, rIndex) => {
    row.forEach((font, cIndex) => {
      const cellAddress = getCellAddress(rIndex + range.getRow(), cIndex + range.getColumn());

      if (font !== "Arial") {
        errors.push(`${cellAddress}: Wrong font family (should be Arial)`);
      }

      const color = colors[rIndex][cIndex];
      if (color !== "#000000") {
        errors.push(`${cellAddress}: Wrong text color (should be black)`);
      }
    });
  });
}

function checkButtons(sheet, errors) {
  const buttonStructure = {
    H: {
      content: ["Clear Data", "Fill Example Data", "Save Snapshot", "Load Last Snapshot", "Convert to Main Currency"],
      format: {
        hAlign: "left",
        vAlign: "middle",
        fontSize: 10,
        style: "normal",
        startRow: 2,
        endRow: 6,
      },
    },
    I: {
      content: [
        "Remove all data but keep headers",
        "Insert sample financial data",
        "Save a copy with UTC timestamp",
        "Restore the last saved snapshot",
        "Fetch exchange rates and recalculate",
      ],
      format: {
        hAlign: "left",
        vAlign: "middle",
        fontSize: 10,
        style: "italic",
        startRow: 2,
        endRow: 6,
      },
    },
    J: {
      content: Array(5).fill(""),
      format: {
        hAlign: "center",
        vAlign: "middle",
        fontSize: 10,
        style: "normal",
        startRow: 2,
        endRow: 6,
        isCheckbox: true,
      },
    },
  };

  Object.entries(buttonStructure).forEach(([col, config]) => {
    const { content, format } = config;

    for (let rowIndex = 0; rowIndex < content.length; rowIndex++) {
      const row = format.startRow + rowIndex;
      const cell = sheet.getRange(`${col}${row}`);
      const expected = content[rowIndex];

      if (expected && cell.getValue() !== expected) {
        errors.push(`${col}${row}: Should contain "${expected}"`);
      }

      if (format.isCheckbox) {
        const validation = cell.getDataValidation();
        if (!validation || validation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
          errors.push(`${col}${row}: Should be a checkbox`);
        }
      }

      checkButtonFormatting(cell, format, errors);
    }
  });
}

function checkButtonFormatting(cell, format, errors) {
  const address = cell.getA1Notation();

  const textStyle = cell.getTextStyle();
  const fontSize = cell.getFontSize();
  const hAlign = cell.getHorizontalAlignment().toLowerCase();
  const vAlign = cell.getVerticalAlignment().toLowerCase();
  if (hAlign !== format.hAlign) {
    errors.push(`${address}: Wrong horizontal alignment (should be ${format.hAlign})`);
  }
  if (vAlign !== format.vAlign) {
    errors.push(`${address}: Wrong vertical alignment (should be ${format.vAlign})`);
  }

  if (fontSize !== format.fontSize) {
    errors.push(`${address}: Wrong font size (should be ${format.fontSize})`);
  }

  const unwantedStyles = [];

  const hasItalic = textStyle.isItalic();
  if (format.style === "italic" && !hasItalic) {
    errors.push(`${address}: Missing italic style`);
  } else if (format.style === "normal" && hasItalic) {
    unwantedStyles.push("italic");
  }

  if (textStyle.isBold()) {
    unwantedStyles.push("bold");
  }

  if (cell.getFontLine() === "line-through") {
    unwantedStyles.push("strikethrough");
  }

  if (cell.getFontLine() === "underline") {
    unwantedStyles.push("underline");
  }

  if (unwantedStyles.length > 0) {
    errors.push(`${address}: Found unwanted styles (${unwantedStyles.join(", ")})`);
  }
}

function getCellAddress(row, col) {
  return `${columnToLetter(col)}${row}`;
}

function columnToLetter(column) {
  let temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function checkBorders(range, errors) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileId = ss.getId();
  const sheetName = range.getSheet().getName();
  const token = ScriptApp.getOAuthToken();
  const rangeA1 = range.getA1Notation();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const [expectedRows, expectedCols] = [range.getNumRows(), range.getNumColumns()];
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${fileId}?ranges=${encodeURIComponent(
    sheetName + "!" + rangeA1
  )}&fields=sheets/data/rowData/values/userEnteredFormat/borders`;

  try {
    const params = {
      method: "get",
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, params);
    const content = response.getContentText();
    const result = JSON.parse(content);

    if (!result.sheets) {
      errors.push(`${rangeA1}: No sheet data`);
      return;
    }

    const sheetData = result.sheets[0];
    if (!sheetData?.data?.[0]?.rowData) {
      errors.push(`${rangeA1}: No row data`);
      return;
    }

    const rowData = sheetData.data[0].rowData;

    if (rowData.length !== expectedRows) {
      errors.push(`${rangeA1}: Missing rows. Expected ${expectedRows} rows, got ${rowData.length}`);
      return;
    }

    rowData.forEach((row, rowIndex) => {
      if (!row.values || row.values.length !== expectedCols) {
        const actualRow = startRow + rowIndex;
        errors.push(`${rangeA1}: Row ${actualRow} has missing columns. Expected ${expectedCols} columns, got ${row.values?.length || 0}`);
        return;
      }

      row.values.forEach((cell, colIndex) => {
        const borders = cell?.userEnteredFormat?.borders;
        const actualRow = startRow + rowIndex;
        const actualCol = startCol + colIndex;
        const cellAddress = `${columnToLetter(actualCol)}${actualRow}`;

        if (!borders) {
          errors.push(`${cellAddress}: No borders defined`);
          return;
        }

        ["top", "bottom", "left", "right"].forEach((side) => {
          if (!borders[side]) {
            errors.push(`${cellAddress}: Missing ${side} border`);
          }
        });

        ["top", "bottom", "left", "right"].forEach((side) => {
          const border = borders[side];
          if (border) {
            if (border.style !== "SOLID") {
              errors.push(`${cellAddress}: ${side} border should be solid`);
            }
            if (border.color && (border.color.red > 0 || border.color.green > 0 || border.color.blue > 0)) {
              errors.push(`${cellAddress}: ${side} border should be black`);
            }
          }
        });
      });
    });

    const lastRow = startRow + expectedRows - 1;
    const lastCol = startCol + expectedCols - 1;
    const expectedRange = `${columnToLetter(startCol)}${startRow}:${columnToLetter(lastCol)}${lastRow}`;

    if (rangeA1 !== expectedRange) {
      errors.push(`Range mismatch: Expected ${expectedRange}, got ${rangeA1}`);
    }
  } catch (error) {
    errors.push(`${rangeA1}: Critical error - ${error.message}`);
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

    Browser.msgBox("✅ Cache cleared successfully!");
  } catch (error) {
    Browser.msgBox("❌ Error clearing cache: " + error.message);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Financial Tools")
    .addItem("Clear Cache", "clearAllCache")
    .addItem("Check Structure", "checkSheetStructure")
    .addItem("Restore Structure", "restoreAllStructure")
    .addToUi();
}

function restoreAllStructure() {
  try {
    restoreInitialStructure();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Financial Summary");
    if (!sheet) {
      throw new Error("Financial Summary sheet not found");
    }

    restoreExternalStructure(sheet);
    restoreInternalStructure(sheet);

    Browser.msgBox("✅ All structures restored successfully!");
  } catch (error) {
    Browser.msgBox("❌ Error during restore: " + error.message);
  }
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
        Browser.msgBox(`❌ Error: ${error.message}`);
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
    Browser.msgBox("❌ Sheet not found!");
    return;
  }

  var range = sheet.getRange("A2:F1000");

  try {
    range.clear();

    range
      .setFontFamily("Arial")
      .setFontSize(10)
      .setFontStyle("normal")
      .setFontWeight("normal")
      .setFontLine("none")
      .setFontColor("#000000")
      .setBackground(null)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("General")
      .setShowHyperlink(false)
      .setTextRotation(0);

    range.setDataValidation(null);

    range.clearNote();

    var rules = sheet.getConditionalFormatRules();
    var newRules = rules.filter((rule) => {
      var ranges = rule.getRanges();
      return !ranges.some((r) => r.getA1Notation().includes("A2:F1000"));
    });
    sheet.setConditionalFormatRules(newRules);

    range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    Browser.msgBox("✅ Sheet cleared and reset successfully!");
  } catch (error) {
    Browser.msgBox("❌ Error: " + error.message);
  }
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

  restoreInternalStructure(sheet);

  Browser.msgBox("✅ Example data inserted!");
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

  Browser.msgBox("✅ Snapshot saved as: " + newSheetName);
}

function loadLastSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss
    .getSheets()
    .map((sheet) => sheet.getName())
    .filter((name) => name.startsWith("Snapshot_"));

  if (sheets.length === 0) {
    Browser.msgBox("❌ No snapshots found!");
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

  Browser.msgBox("✅ Restored from: " + lastSnapshotName);
}

function onError(error) {
  Browser.msgBox(`❌ Critical error: ${error.message}`);
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
  Browser.msgBox("❌ TOTAL: row not found");
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

  for (let i = 1; i < totalRowIndex; i++) {
    const rowLabel = data[i][0]?.toString().trim().toUpperCase() || "";

    if (rowLabel.startsWith("SUBTOTAL:")) {
      if (currentSection.length > 0) {
        data[i][1] = `=SUM(E${currentSection[0]}:E${currentSection[currentSection.length - 1]})`;
      }
      if (data[i][2]?.toString().toUpperCase() !== mainCurrency) {
        data[i][2] = mainCurrency;
      }
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
  Browser.msgBox(`✅ Converted to ${mainCurrency}`);
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
    Browser.msgBox("❌ Error: " + e.message);
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
      Browser.msgBox(`❌ Error: ${error.message}`);
    }

    e.range.setValue(false);
  }
}
