/**
 * Service_Materials.gs
 */
function getMaterialsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Materials");
    const manageSheet = ss.getSheetByName("Manage");
    
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    
    // 1. If no project is selected, return an empty array immediately
    if (!rowIndex || !manageSheet) return [];

    const selectedPid = manageSheet.getRange(Number(rowIndex), 1).getValue();
    if (!selectedPid) return [];

    // 2. Safety check for the Materials sheet existence
    if (!sheet) {
      console.error("Materials sheet not found");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    // 3. Map with updated columns (A through G)
    return data.slice(1)
      .map((row, index) => {
        return {
          pid: row[0] ? String(row[0]) : "",       // Col A (1)
          item: row[1] ? String(row[1]) : "",      // Col B (2)
          qty: row[2] || 0,                        // Col C (3)
          taskName: row[3] ? String(row[3]) : "",  // Col D (4)
          taskId: row[4] ? String(row[4]) : "",    // Col E (5)
          roomName: row[5] ? String(row[5]) : "",  // Col F (6) - NEW
          roomId: row[6] ? String(row[6]) : "",    // Col G (7) - NEW
          sheetRow: index + 2
        };
      })
      .filter(m => String(m.pid) === String(selectedPid));
      
  } catch (e) {
    console.error("Error in getMaterialsData: " + e.toString());
    return []; // Return empty array so the UI can at least stop spinning
  }
}

function saveMaterialsData(materObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Materials");
  
  // MAPPING (matches Equipment structure): 
  // Col 2: Item (B)
  // Col 3: Qty (C)
  // Col 4: TaskName (D)
  // Col 5: TaskID (E)
  // Col 6: RoomName (F)
  // Col 7: RoomID (G)
  const rowValues = [[
    String(materObj.item), 
    Number(materObj.qty), 
    String(materObj.taskName), 
    String(materObj.taskId),
    String(materObj.roomName),
    String(materObj.roomId)
  ]];

  if (materObj.sheetRow && Number(materObj.sheetRow) > 0) {
    // Update starting at Col B (2), spanning 6 columns
    sheet.getRange(Number(materObj.sheetRow), 2, 1, 6).setValues(rowValues);
  } else {
    const userProperties = PropertiesService.getUserProperties();
    const manageRow = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    const pid = ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue();
    
    // Append [PiD, Item, Qty, TaskName, TaskID, RoomName, RoomID]
    sheet.appendRow([pid, ...rowValues[0]]);
  }
  return getMaterialsData();
}

/**
 * DELETE MATERIAL
 * Removes the specific row and returns updated data.
 */
function deleteMaterial(rowIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Materials");
    const rowToDelete = parseInt(rowIdx, 10);

    // Safety check: ensure sheet exists and row is valid (not header)
    if (!sheet) throw new Error("Materials sheet not found.");
    if (isNaN(rowToDelete) || rowToDelete < 2 || rowToDelete > sheet.getLastRow()) {
      throw new Error("Invalid row index: " + rowIdx);
    }

    sheet.deleteRow(rowToDelete);
    
    // Return fresh list so UI updates immediately
    return getMaterialsData();
  } catch (e) {
    throw new Error("Delete failed: " + e.toString());
  }
}