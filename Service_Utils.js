/**
 * UNIVERSAL FETCHER: Enhanced with Pre-Flight Check
 */
function getFilteredData(sheetName, pidColIdx) {
  const activePid = getActivePid();
  if (!activePid) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  // PRE-FLIGHT GUARD
  if (!sheet || sheet.getLastRow() === 0) return [];

  // If there is only 1 row (the header), return empty
  if (sheet.getLastRow() === 1) return [];

  const data = sheet.getDataRange().getValues();

  return data.slice(1)
    .map((row, index) => ({ data: row, rowIdx: index + 2 }))
    .filter(obj => String(obj.data[pidColIdx]) === String(activePid));
}

/**
 * Service_Utils.gs
 */
function getActivePid() {
  const props = PropertiesService.getUserProperties();
  
  // 1. Primary: Try direct PID
  const pid = props.getProperty('ACTIVE_PID');
  if (pid) return pid.trim();

  // 2. Fallback: Try Row Index
  const rowIndex = props.getProperty('ACTIVE_PROJECT_ROW');
  if (!rowIndex) return null; // No need to throw if we can just return null

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Manage");
    if (!sheet) return null;

    return sheet.getRange(parseInt(rowIndex), 1).getDisplayValue().trim();
  } catch (e) {
    console.error("Error retrieving PID from row: " + e.message);
    return null;
  }
}

/**
 * Generic helper to save data to a specific sheet and row.
 */
function saveCommonData(sheetName, rowIndex, dataArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) throw new Error("Sheet '" + sheetName + "' not found.");

  const targetRow = Number(rowIndex);
  if (isNaN(targetRow) || targetRow < 1) {
    throw new Error("Invalid row index provided.");
  }

  // Updates starting from Column A (1), for 1 row, across the length of the data
  sheet.getRange(targetRow, 1, 1, dataArray.length).setValues([dataArray]);
  
  return true;
}
