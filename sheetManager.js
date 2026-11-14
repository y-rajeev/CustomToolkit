// -----------------
// Sheet Manipulator v2
// -----------------

const SHEET_MANIPULATOR_NAME = "Sheet Manipulator";
const PROTECTED_SHEETS = [SHEET_MANIPULATOR_NAME];

function listSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MANIPULATOR_NAME);
  const outputRange = sheet.getRange("A4:J");
  outputRange.clearContent();

  const sheets = ss.getSheets();
  const out = sheets.map(s => [s.getName(), false, false, '', false, '', false, '', '', '']);
  sheet.getRange(4, 1, out.length, 10).setValues(out);
}

function deleteSheets() {
  manipulateSheets({
    conditionIndex: 1,
    action: (ss, sheetName) => {
      if (PROTECTED_SHEETS.includes(sheetName)) throw new Error(`Cannot delete protected sheet: ${sheetName}`);
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
      ss.deleteSheet(sheet);
    },
    resultText: name => `Deleted ${name}`
  });
}

function copySheets() {
  manipulateSheets({
    conditionIndex: 2,
    action: (ss, sheetName, row) => {
      const targetId = row[3];
      if (!targetId) throw new Error("Missing destination Sheet ID in Column D.");
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
      const targetSS = SpreadsheetApp.openById(targetId);
      const copied = sheet.copyTo(targetSS);
      copied.setName(sheetName + " - Copy " + new Date().toISOString().slice(0, 10));
    },
    resultText: name => `Copied ${name}`
  });
}

function moveSheets() {
  manipulateSheets({
    conditionIndex: 4,
    action: (ss, sheetName, row) => {
      const targetId = row[5];
      if (!targetId) throw new Error("Missing destination Sheet ID in Column F.");
      if (PROTECTED_SHEETS.includes(sheetName)) throw new Error(`Cannot delete protected sheet: ${sheetName}`);
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
      const targetSS = SpreadsheetApp.openById(targetId);
      const copied = sheet.copyTo(targetSS);
      copied.setName(sheetName + " - Moved " + new Date().toISOString().slice(0, 10));
      ss.deleteSheet(sheet);
    },
    resultText: name => `Moved ${name}`
  });
}

function renameSheets() {
  manipulateSheets({
    conditionIndex: 6,
    action: (ss, sheetName, row) => {
      const newName = row[7];
      if (!newName) throw new Error("Missing new name in Column H.");
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
      if (PROTECTED_SHEETS.includes(sheetName)) throw new Error(`Cannot rename protected sheet: ${sheetName}`);
      sheet.setName(newName);
    },
    resultText: (name, row) => `Renamed ${name} to ${row[7]}`
  });
}

function manipulateSheets({ conditionIndex, action, resultText }) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MANIPULATOR_NAME);
  const range = sheet.getRange("A4:J");
  const values = range.getValues();
  const rowCount = values.reduce((lastIndex, row, index) => row[0] ? index + 1 : lastIndex, 0);

  if (rowCount === 0) return;

  for (let i = 0; i < rowCount; i++) {
    const row = values[i];
    const sheetName = row[0];
    const condition = row[conditionIndex];

    if (condition === true) {
      try {
        action(ss, sheetName, row);
        values[i][8] = typeof resultText === 'function'
          ? resultText(sheetName, row)
          : resultText(sheetName);
      } catch (e) {
        values[i][8] = `Error: ${e.toString()}`;
        Logger.log(`Error on row ${i + 4}: ${e.toString()}`);
      }
      values[i][9] = new Date(); // Timestamp
      values[i][conditionIndex] = false; // Uncheck after execution
    }
  }

  sheet.getRange(4, 1, rowCount, 10).setValues(values.slice(0, rowCount));
}
