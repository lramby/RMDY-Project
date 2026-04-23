/**
 * Service_Tasks.gs
 */

function getTasksData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    if (!rowIndex) return [];

    const manageSheet = ss.getSheetByName("Manage");
    const selectedPid = manageSheet.getRange(Number(rowIndex), 1).getValue();

    // USE CONFIG: Sheet Name
    const taskSheet = ss.getSheetByName(CONFIG.TABLES.TASKS.NAME);
    const data = taskSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    return data.slice(1).map((row, index) => {

      const COLS = CONFIG.TABLES.TASKS.COLUMNS;

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
        taskName: row[COLS.TASKNAME] ? String(row[COLS.TASKNAME]) : "",
        value: row[COLS.VALUE] || 0,
        unit: row[COLS.UNIT] ? String(row[COLS.UNIT]) : "",
        roomName: row[COLS.ROOMNAME] || "No Room Assigned",
        roomId: row[COLS.ROOMID] ? String(row[COLS.ROOMID]) : "",
        taskId: row[COLS.TASKID] ? String(row[COLS.TASKID]) : "",
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

function saveTask(taskObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABLES.TASKS.NAME);
  const COLS = CONFIG.TABLES.TASKS.COLUMNS;
  
  // Create array based on column length defined in your sheet schema
  const rowArray = new Array(10).fill("");

  // Map taskObj values using CONFIG column indices 
  rowArray[COLS.ROOMID]   = String(taskObj.roomId || "");
  rowArray[COLS.TASKID]   = String(taskObj.taskId || "");
  rowArray[COLS.ROOMNAME] = String(taskObj.roomName || "");
  rowArray[COLS.TASKNAME] = String(taskObj.taskName || "");
  rowArray[COLS.VALUE]    = Number(taskObj.value) || 0;
  rowArray[COLS.UNIT]     = String(taskObj.unit || "");
  rowArray[COLS.PRICE]    = Number(taskObj.price) || 0;
  rowArray[COLS.COST]     = Number(taskObj.cost) || 0;
  rowArray[COLS.NOTE]     = String(taskObj.note || "");

  if (taskObj.sheetRow && Number(taskObj.sheetRow) >= 2) {
    // UPDATE: Skip PID (Index 0) to maintain project association
    const updateRange = rowArray.slice(1); 
    sheet.getRange(Number(taskObj.sheetRow), 2, 1, updateRange.length).setValues([updateRange]);
  } else {
    // NEW: Retrieve current Project ID for new task entries
    const userProperties = PropertiesService.getUserProperties();
    const manageRow = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    const pid = ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue();
    
    rowArray[COLS.PID] = pid;
    sheet.appendRow(rowArray);
  }
  
  return getTasksData();
}

function deleteTask(rowIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TABLES.TASKS.NAME);
    const rowToDelete = parseInt(rowIdx, 10);

    if (isNaN(rowToDelete) || rowToDelete < 2) {
      throw new Error("Invalid row index.");
    }

    sheet.deleteRow(rowToDelete);
    return getTasksData();
  } catch (e) {
    console.error("Delete Task failed: " + e.toString());
    throw new Error(e.toString());
  }
}





/**======================================
* Room Backend Functions
*========================================*/

function getRoomOptionsForTask() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    if (!rowIndex) return [];

    // Force PID to String
    const pid = String(ss.getSheetByName("Manage").getRange(Number(rowIndex), 1).getValue());
    
    // Using the new sheet name
    const roomSheet = ss.getSheetByName("Rooms");
    if (!roomSheet) return [];
    
    const roomData = roomSheet.getDataRange().getValues();

    return roomData.slice(1)
      .filter(row => String(row[0]) === pid) // Forced string comparison makes it type-safe
      .map(row => ({
        id: String(row[6]), // RoomID from Column G
        display: row[2] ? `${row[1]} (#${row[2]})` : row[1]
      }));
  } catch (e) {
    console.error("Room Dropdown Error: " + e.toString());
    return [];
  }
}

/**
 * Updates linked materials/equipment when a task's room info changes.
 */
function cascadeRoomNameUpdate(roomId, newRoomName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToUpdate = ["Materials", "Equipment"];
  
  sheetsToUpdate.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      // If Room ID (Col F / Index 5) matches, update Room Name (Col D / Index 3)
      if (String(data[i][5]) === String(roomId)) {
        sheet.getRange(i + 1, 4).setValue(newRoomName);
      }
    }
  });
}

/**
 * Simply performs the row deletion safely.
 */
function safeDeleteRow(sheetName, rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet && rowIdx > 1) {
    sheet.deleteRow(rowIdx);
  }
}