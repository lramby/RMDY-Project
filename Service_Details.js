/**
 * Service_Details.gs
 */

function getDetailsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    
    if (!rowIndex) return null;

    // USE CONFIG: Sheet Name
    const sheet = ss.getSheetByName(CONFIG.TABLES.DETAILS.NAME);
    const targetRow = Number(rowIndex);
    
    if (targetRow < 2 || targetRow > sheet.getLastRow()) return null;

    // USE CONFIG: Columns
    const COLS = CONFIG.TABLES.DETAILS.COLUMNS;
    const rowData = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    return {
      pid:       rowData[COLS.PID] ? String(rowData[COLS.PID]) : "",
      address:   rowData[COLS.ADDRESS] || "",
      address2:  rowData[COLS.ADDRESS2] || "",
      city:      rowData[COLS.CITY] || "",
      zip:       rowData[COLS.ZIP] || "",
      state:     rowData[COLS.STATE] || "",
      country:   rowData[COLS.COUNTRY] || "",
      firstName: rowData[COLS.FIRSTNAME] || "",
      lastName:  rowData[COLS.LASTNAME] || "",
      email:     rowData[COLS.EMAIL] || "",
      phone:     rowData[COLS.PHONE] || "",
      sheetRow:  targetRow
    };
  } catch (e) {
    console.error("Error in getDetailsData: " + e.toString());
    return null;
  }
}

function updateProjectDetails(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABLES.DETAILS.NAME);
  const COLS = CONFIG.TABLES.DETAILS.COLUMNS;
  
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');

  if (!rowIndex) throw new Error("No active project selected.");
  const targetRow = Number(rowIndex);
  
  // 1. Create an array representing the row (adjust length based on your schema)
  // We initialize with empty strings
  const rowArray = [];
  
  // 2. Map formObject values to exact positions defined in CONFIG
  // Note: We use specific getRange calls or a full array to match the pattern
  // To keep it clean and match the Service_Equipment pattern of skipping PID (Col A):
  
  const updateData = [];
  // We update from Column B (Index 1) onwards to avoid overwriting PID
  updateData[COLS.ADDRESS - 1]   = formObject.address || "";
  updateData[COLS.ADDRESS2 - 1]  = formObject.address2 || "";
  updateData[COLS.CITY - 1]      = formObject.city || "";
  updateData[COLS.ZIP - 1]       = formObject.zip || "";
  updateData[COLS.STATE - 1]     = formObject.state || "";
  updateData[COLS.COUNTRY - 1]   = formObject.country || "";
  updateData[COLS.FIRSTNAME - 1] = formObject.firstName || "";
  updateData[COLS.LASTNAME - 1]  = formObject.lastName || "";
  updateData[COLS.EMAIL - 1]     = formObject.email || "";
  updateData[COLS.PHONE - 1]     = formObject.phone || "";

  // Set the values starting from Column B (Index 2)
  sheet.getRange(targetRow, 2, 1, updateData.length).setValues([updateData]);

  return getDetailsData();
}