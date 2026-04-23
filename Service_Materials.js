/**
 * Service_Materials.gs
 * Refactored to use CONFIG object and updated column mapping.
 */

function getMaterialsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TABLES.MATERIALS.NAME);
    const manageSheet = ss.getSheetByName("Manage");
    const cols = CONFIG.TABLES.MATERIALS.COLUMNS;
    
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    
    // 1. If no project is selected, return an empty array
    if (!rowIndex || !manageSheet) return [];

    const selectedPid = manageSheet.getRange(Number(rowIndex), 1).getValue();
    if (!selectedPid) return [];

    // 2. Safety check for the Materials sheet existence
    if (!sheet) {
      console.error("Materials sheet not found: " + CONFIG.TABLES.MATERIALS.NAME);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    // 3. Map using CONFIG indices
    return data.slice(1)
      .map((row, index) => {
        return {
          pid: row[cols.PID] ? String(row[cols.PID]) : "",
          roomId: row[cols.ROOMID] ? String(row[cols.ROOMID]) : "",
          taskId: row[cols.TASKID] ? String(row[cols.TASKID]) : "",
          roomName: row[cols.ROOMNAME] ? String(row[cols.ROOMNAME]) : "",
          taskName: row[cols.TASKNAME] ? String(row[cols.TASKNAME]) : "",
          item: row[cols.ITEMNAME] ? String(row[cols.ITEMNAME]) : "",
          value: row[cols.VALUE] || 0, // Changed from qty to value
          unit: row[cols.UNIT] || "",
          price: row[cols.PRICE] || 0,
          cost: row[cols.COST] || 0,
          note: row[cols.NOTE] || "",
          sheetRow: index + 2
        };
      })
      .filter(m => String(m.pid) === String(selectedPid));
      
  } catch (e) {
    console.error("Error in getMaterialsData: " + e.toString());
    return [];
  }
}

function saveMaterialsData(materObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABLES.MATERIALS.NAME);
  const cols = CONFIG.TABLES.MATERIALS.COLUMNS;
  
  // Create an array the size of our column definitions to ensure correct placement
  // We use the indices from CONFIG to place data in the correct slots
  const rowValues = [];
  rowValues[cols.ROOMID] = String(materObj.roomId);
  rowValues[cols.TASKID] = String(materObj.taskId);
  rowValues[cols.ROOMNAME] = String(materObj.roomName);
  rowValues[cols.TASKNAME] = String(materObj.taskName);
  rowValues[cols.ITEMNAME] = String(materObj.item);
  rowValues[cols.VALUE] = Number(materObj.value); // Maps to the VALUE column
  rowValues[cols.UNIT] = String(materObj.unit || "");
  rowValues[cols.PRICE] = Number(materObj.price || 0);
  rowValues[cols.COST] = Number(materObj.cost || 0);
  rowValues[cols.NOTE] = String(materObj.note || "");

  if (materObj.sheetRow && Number(materObj.sheetRow) > 0) {
    // Update existing row: Start at index 1 (Col B) because PID (index 0) usually stays the same
    // We calculate the length minus the PID column
    const updateArray = rowValues.slice(1); 
    sheet.getRange(Number(materObj.sheetRow), 2, 1, updateArray.length).setValues([updateArray]);
  } else {
    // New entry: Get Active PID
    const userProperties = PropertiesService.getUserProperties();
    const manageRow = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    const pid = ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue();
    
    rowValues[cols.PID] = pid;
    sheet.appendRow(rowValues);
  }
  return getMaterialsData();
}

function deleteMaterial(rowIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TABLES.MATERIALS.NAME);
    const rowToDelete = parseInt(rowIdx, 10);

    if (!sheet) throw new Error("Materials sheet not found.");
    if (isNaN(rowToDelete) || rowToDelete < 2) throw new Error("Invalid row index.");

    sheet.deleteRow(rowToDelete);
    return getMaterialsData();
  } catch (e) {
    throw new Error("Delete failed: " + e.toString());
  }
}