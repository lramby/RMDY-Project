/**
 * Service_Dates.gs
 */

function getDatesData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TABLES.DATES.NAME); // Use CONFIG for sheet name
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    
    if (!rowIndex) return null;
    const targetRow = parseInt(rowIndex);
    const COLS = CONFIG.TABLES.DATES.COLUMNS;

    // Fetch the specific project row
    // We fetch 12 columns as defined in your current setup
    const displayValues = sheet.getRange(targetRow, 1, 1, 12).getDisplayValues()[0];
    
    return {
      pid:       displayValues[COLS.PID],
      loss:      displayValues[COLS.LOSS],
      due:       displayValues[COLS.DUE],
      contacted: displayValues[COLS.CONTACTED],
      assigned:  displayValues[COLS.ASSIGNED],
      inspected: displayValues[COLS.INSPECTED],
      estimated: displayValues[COLS.ESTIMATED],
      started:   displayValues[COLS.STARTED],
      finished:  displayValues[COLS.FINISHED],
      invoiced:  displayValues[COLS.INVOICED],
      approved:  displayValues[COLS.APPROVED],
      paid:      displayValues[COLS.PAID]
    };
  } catch (e) {
    console.error("Error in getDatesData: " + e.toString());
    return null;
  }
}

function updateDatesData(formObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TABLES.DATES.NAME);
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');

    if (!rowIndex) throw new Error("No active project selected.");
    const targetRow = parseInt(rowIndex);
    const COLS = CONFIG.TABLES.DATES.COLUMNS;
    
    // 1. Create a full row array to ensure data lands in the right spots
    // Based on your original code, you are updating 11 columns (B-L)
    const rowArray = new Array(12).fill(""); 

    // 2. Map formObject values to the specific CONFIG indices
    rowArray[COLS.LOSS]      = formObject.loss;
    rowArray[COLS.DUE]       = formObject.due;
    rowArray[COLS.CONTACTED] = formObject.contacted;
    rowArray[COLS.ASSIGNED]  = formObject.assigned;
    rowArray[COLS.INSPECTED] = formObject.inspected;
    rowArray[COLS.ESTIMATED] = formObject.estimated;
    rowArray[COLS.STARTED]   = formObject.started;
    rowArray[COLS.FINISHED]  = formObject.finished;
    rowArray[COLS.INVOICED]  = formObject.invoiced;
    rowArray[COLS.APPROVED]  = formObject.approved;
    rowArray[COLS.PAID]      = formObject.paid;

    // 3. Write back to the sheet (Columns 2 through 12, which is index 1 to 11)
    // We skip PID (Col A) to prevent accidental overwriting of the unique ID
    const updateValues = [rowArray.slice(1)]; 
    sheet.getRange(targetRow, 2, 1, 11).setValues(updateValues);

    return getDatesData();
  } catch (e) {
    console.error("Error in updateDatesData: " + e.toString());
    throw e;
  }
}