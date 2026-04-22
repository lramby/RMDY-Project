/**
 * Service_Service.gs
 * Fetches and Updates Service-specific data
 */

function getServiceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Service");
  
  // 1. Get the row index saved for THIS specific user
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  
  // 2. Return null if no project is selected
  if (!rowIndex) return null;

  const targetRow = parseInt(rowIndex);
  const lastRow = sheet.getLastRow();

  // 3. Safety check for row validity
  if (targetRow > lastRow || targetRow < 2) return null;

  // 4. Fetch ONLY the specific row data (Columns A through I)
  // getRange(row, column, numRows, numColumns)
  const displayValues = sheet.getRange(targetRow, 1, 1, 9).getDisplayValues()[0];
  
  return {
    pid:          displayValues[0], // Col A
    serviceType:  displayValues[1], // Col B
    setup:         displayValues[2], // Col C
    takedown:      displayValues[3], // Col D
    monitor:       displayValues[4], // Col E
    drivetime:     displayValues[5], // Col F
    thermal:       displayValues[6], // Col G
    fog:           displayValues[7], // Col H
    description:   displayValues[8]  // Col I
  };
}

function updateServiceData(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Service");
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');

  if (!rowIndex) throw new Error("No active project selected.");
  const targetRow = parseInt(rowIndex);
  
  // Updating Col B through I (8 columns)
  const values = [[
    formObject.serviceType, 
    formObject.setup, 
    formObject.takedown, 
    formObject.monitor, 
    formObject.drivetime, 
    formObject.thermal, 
    formObject.fog, 
    formObject.description
  ]];
  
  sheet.getRange(targetRow, 2, 1, 8).setValues(values);
  
  // This calls the function above to return the fresh data back to the UI
  return getServiceData();
}