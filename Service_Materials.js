/**
 * Service_Materials.gs
 * Refactored to include price/cost subtotals and mirror Equipment structure.
 */

function getMaterialsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    if (!rowIndex) return [];

    const manageSheet = ss.getSheetByName("Manage");
    const selectedPid = manageSheet.getRange(Number(rowIndex), 1).getValue();
    
    // USE CONFIG: Sheet Name
    const materialSheet = ss.getSheetByName(CONFIG.TABLES.MATERIALS.NAME);
    const data = materialSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    return data.slice(1).map((row, index) => {
        const COLS = CONFIG.TABLES.MATERIALS.COLUMNS;

        // 1. Extract raw numbers for calculation
        const rawValue = Number(row[COLS.VALUE]) || 0;
        const rawPrice = Number(row[COLS.PRICE]) || 0;
        const rawCost  = Number(row[COLS.COST]) || 0;

        // 2. Perform the math
        const rawPriceSubTotal = rawValue * rawPrice;
        const rawCostSubTotal = rawValue * rawCost;

        // 3. Return the object with formatted strings
        return {
            pid: row[COLS.PID] ? String(row[COLS.PID]) : "",
            roomId: row[COLS.ROOMID] ? String(row[COLS.ROOMID]) : "",
            taskId: row[COLS.TASKID] ? String(row[COLS.TASKID]) : "",
            roomName: row[COLS.ROOMNAME] || "No Room Assigned",
            taskName: row[COLS.TASKNAME] || "No Task Assigned",
            item: row[COLS.ITEMNAME] ? String(row[COLS.ITEMNAME]) : "",
            value: rawValue,
            unit: row[COLS.UNIT] ? String(row[COLS.UNIT]) : "",
            note: row[COLS.NOTE] ? String(row[COLS.NOTE]) : "-",

            // Formatting via Config
            price: CONFIG.FORMAT.CURRENCY(rawPrice),
            cost: CONFIG.FORMAT.CURRENCY(rawCost),
            priceSubTotal: CONFIG.FORMAT.CURRENCY(rawPriceSubTotal),
            costSubTotal: CONFIG.FORMAT.CURRENCY(rawCostSubTotal),

            sheetRow: index + 2 
        };
    }).filter(mater => String(mater.pid) === String(selectedPid));
    
  } catch (e) {
    console.error("Error in getMaterialsData: " + e.toString());
    return [];
  }
}

function saveMaterialsData(materObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABLES.MATERIALS.NAME);
  const COLS = CONFIG.TABLES.MATERIALS.COLUMNS;
  
  // 1. Create an array representing the full row (matches Materials column count)
  // Initializing with empty strings to prevent 'undefined' in cells
  const rowArray = new Array(11).fill("");

  // 2. Map the materObj values to the exact positions defined in CONFIG
  rowArray[COLS.ROOMID]   = String(materObj.roomId || "");
  rowArray[COLS.TASKID]   = String(materObj.taskId || "");
  rowArray[COLS.ROOMNAME] = String(materObj.roomName || "");
  rowArray[COLS.TASKNAME] = String(materObj.taskName || "");
  rowArray[COLS.ITEMNAME] = String(materObj.item || "");
  rowArray[COLS.VALUE]    = Number(materObj.value) || 0;
  rowArray[COLS.UNIT]     = String(materObj.unit || "");
  rowArray[COLS.PRICE]    = Number(materObj.price) || 0;
  rowArray[COLS.COST]     = Number(materObj.cost) || 0;
  rowArray[COLS.NOTE]     = String(materObj.note || "");

  if (materObj.sheetRow && Number(materObj.sheetRow) >= 2) {
    // UPDATE: Skip PID (Index 0), update from Col B (Index 1) onwards
    const updateRange = rowArray.slice(1); 
    sheet.getRange(Number(materObj.sheetRow), 2, 1, updateRange.length).setValues([updateRange]);
  } else {
    // NEW: Get the PID from the active project row
    const userProperties = PropertiesService.getUserProperties();
    const manageRow = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    const pid = ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue();
    
    rowArray[COLS.PID] = pid;
    sheet.appendRow(rowArray);
  }
  
  return getMaterialsData();
}

function deleteMaterial(rowIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TABLES.MATERIALS.NAME);
    const rowToDelete = parseInt(rowIdx, 10);

    if (isNaN(rowToDelete) || rowToDelete < 2) {
      throw new Error("Invalid row index.");
    }

    sheet.deleteRow(rowToDelete);
    return getMaterialsData();
  } catch (e) {
    console.error("Delete failed: " + e.toString());
    throw new Error(e.toString());
  }
}