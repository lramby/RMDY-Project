/**
 * Service_Details.gs
 */
function getDetailsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Details");
  
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  
  if (!rowIndex) return null;

  const targetRow = parseInt(rowIndex);
  if (targetRow > sheet.getLastRow()) return null;

  // Fetch Columns A through K (11 columns)
  const displayValues = sheet.getRange(targetRow, 1, 1, 11).getDisplayValues()[0];
  
  return {
    pid:       displayValues[0], 
    address:   displayValues[1], 
    address2:  displayValues[2], 
    city:      displayValues[3], 
    zip:       displayValues[4], 
    state:     displayValues[5], 
    country:   displayValues[6], 
    firstName: displayValues[7], 
    lastName:  displayValues[8], 
    email:     displayValues[9],  
    phone:     displayValues[10]  
  };
}

function updateProjectDetails(formObject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Details");
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');

  if (!rowIndex) throw new Error("No active project selected.");
  const targetRow = parseInt(rowIndex);
  
  sheet.getRange(targetRow, 2, 1, 6).setValues([[
    formObject.address, 
    formObject.address2, 
    formObject.city, 
    formObject.zip, 
    formObject.state, 
    formObject.country
  ]]);

  sheet.getRange(targetRow, 8, 1, 4).setValues([[
    formObject.firstName, 
    formObject.lastName, 
    formObject.email,
    formObject.phone
  ]]);

  return getDetailsData(); // UPDATED to match your new function name
}