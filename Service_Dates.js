/**
 * Service_Dates.gs
 */
/**
 * Service_Dates.gs
 */
function getDatesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dates");
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  
  if (!rowIndex) return null;
  const targetRow = parseInt(rowIndex);

  // Fetching Columns A through L (12 columns)
  const displayValues = sheet.getRange(targetRow, 1, 1, 12).getDisplayValues()[0];
  
  return {
    pid:       displayValues[0],  // Col A
    loss:      displayValues[1],  // Col B
    due:       displayValues[2],  // Col C
    contacted: displayValues[3],  // Col D
    assigned:  displayValues[4],  // Col E
    inspected: displayValues[5],  // Col F
    estimated: displayValues[6],  // Col G
    started:   displayValues[7],  // Col H
    finished:  displayValues[8],  // Col I
    invoiced:  displayValues[9],  // Col J
    approved:  displayValues[10], // Col K
    paid:      displayValues[11]  // Col L
  };
}

function updateDatesData(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dates");
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');

  if (!rowIndex) throw new Error("No active project selected.");
  const targetRow = parseInt(rowIndex);
  
  // Updating Col B through L (Columns 2 to 12)
  const values = [[
    formObject.loss, formObject.due, formObject.contacted, 
    formObject.assigned, formObject.inspected, formObject.estimated, 
    formObject.started, formObject.finished, formObject.invoiced, 
    formObject.approved, formObject.paid
  ]];
  
  sheet.getRange(targetRow, 2, 1, 11).setValues(values);
  return getDatesData();
}