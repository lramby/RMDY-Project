/**
 * Service_Equipment.gs
 */
 
function getEquipmentData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    if (!rowIndex) return [];

    const manageSheet = ss.getSheetByName("Manage");
    const selectedPid = manageSheet.getRange(Number(rowIndex), 1).getValue();
    
    // USE CONFIG: Sheet Name
    const equipSheet = ss.getSheetByName(CONFIG.TABLES.EQUIPMENT.NAME);
    const data = equipSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    return data.slice(1).map((row, index) => {

		const COLS = CONFIG.TABLES.EQUIPMENT.COLUMNS;

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
				item: row[COLS.ITEM] ? String(row[COLS.ITEM]) : "",
				value: row[COLS.VALUE] || 0,
				unit: row[COLS.UNIT] ? String(row[COLS.UNIT]) : "",
				note: row[COLS.NOTE] ? String(row[COLS.NOTE]) : "-",

				// Formatting via Config
				price: CONFIG.FORMAT.CURRENCY(rawPrice),
				cost: CONFIG.FORMAT.CURRENCY(rawCost),
				priceSubTotal: CONFIG.FORMAT.CURRENCY(rawPriceSubTotal),
				costSubTotal: CONFIG.FORMAT.CURRENCY(rawCostSubTotal),

				sheetRow: index + 2 
			};
    }).filter(task => String(task.pid) === String(selectedPid));
  } catch (e) {
    console.error("Error in getTasksData: " + e.toString());
    return [];
  }
}


function saveEquipmentData(equipObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABLES.EQUIPMENT.NAME);
  const COLS = CONFIG.TABLES.EQUIPMENT.COLUMNS;
  
  // 1. Create an array representing the full row (12 columns)
  // We initialize it with empty strings to avoid 'undefined' in cells
  const rowArray = new Array(12).fill("");

  // 2. Map the equipObj values to the exact positions defined in CONFIG
  rowArray[COLS.ITEMID]   = equipObj.itemID || "E-" + new Date().getTime();
  rowArray[COLS.ROOMID]   = String(equipObj.roomId || "");
  rowArray[COLS.TASKID]   = String(equipObj.taskId || "");
  rowArray[COLS.ROOMNAME] = String(equipObj.roomName || "");
  rowArray[COLS.TASKNAME] = String(equipObj.taskName || "");
  rowArray[COLS.ITEM]     = String(equipObj.item || "");
  rowArray[COLS.VALUE]    = Number(equipObj.value) || 0;
  rowArray[COLS.UNIT]     = String(equipObj.unit || "");
  rowArray[COLS.PRICE]    = Number(equipObj.price) || 0;
  rowArray[COLS.COST]     = Number(equipObj.cost) || 0;
  rowArray[COLS.NOTE]     = String(equipObj.note || "");

  if (equipObj.sheetRow && Number(equipObj.sheetRow) >= 2) {
    // UPDATE: We update starting from Column B (Index 1) to end (Index 11)
    // We skip PID (Index 0) so we don't overwrite the project association
    const updateRange = rowArray.slice(1); 
    sheet.getRange(Number(equipObj.sheetRow), 2, 1, updateRange.length).setValues([updateRange]);
  } else {
    // NEW: Get the PID from the active project row
    const userProperties = PropertiesService.getUserProperties();
    const manageRow = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    const pid = ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue();
    
    rowArray[COLS.PID] = pid;
    sheet.appendRow(rowArray);
  }
  
  return getEquipmentData();
}

function deleteEquipment(rowIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Use CONFIG sheet name
    const sheet = ss.getSheetByName(CONFIG.TABLES.EQUIPMENT.NAME);
    const rowToDelete = parseInt(rowIdx, 10);

    if (isNaN(rowToDelete) || rowToDelete < 2) {
      throw new Error("Invalid row index.");
    }

    sheet.deleteRow(rowToDelete);
    return getEquipmentData();
  } catch (e) {
    console.error("Delete failed: " + e.toString());
    throw new Error(e.toString());
  }
}