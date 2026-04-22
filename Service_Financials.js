/**
 * Service_Financials.gs
 * Fetches financial data for the user's privately selected project.
 */
function getFinancialsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Financials");
  
  if (!sheet) {
    console.error("Sheet 'Financials' not found.");
    return null;
  }

  // 1. Get the row index saved for THIS specific user
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  
  // 2. Return null if no project is selected
  if (!rowIndex) return null;

  const targetRow = parseInt(rowIndex);
  const lastRow = sheet.getLastRow();

  // 3. Safety check for row validity
  if (targetRow > lastRow || targetRow < 2) return null;

  // 4. Fetch ONLY the specific row data (Columns A through G)
  const displayValues = sheet.getRange(targetRow, 1, 1, 12).getDisplayValues()[0];
  
  // Helper to safely clean currency strings for calculation
  const toNum = (val) => {
    if (!val) return 0;
    const cleaned = String(val).replace(/[$,\s]/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
  };

  const data = {
    pid:       displayValues[0], // Column A
    materials: displayValues[1], // Column B
    labor:     displayValues[2], // Column C
    subtrade:  displayValues[3], // Column D
    equipment: displayValues[4], // Column E
    expense:   displayValues[5], // Column F
    overhead:  displayValues[6],  // Column G
    total:  displayValues[7],  // Column G
    invoiced:  displayValues[8],  // Column G
    invdelta:  displayValues[9],  // Column G
    estimated:  displayValues[10],  // Column G
    estdelta:  displayValues[11],  // Column G
  };

  // 5. Calculate the total sum for the Web App display
  data.totalCost = toNum(data.materials) + 
                   toNum(data.labor) + 
                   toNum(data.subtrade) + 
                   toNum(data.equipment) + 
                   toNum(data.expense) + 
                   toNum(data.overhead);
  
  return data;
}