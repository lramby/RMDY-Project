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
        taskName: row[COLS.TASK] ? String(row[COLS.TASK]) : "",
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
  const rowArray = new Array(11).fill("");

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  if (!rowIndex) return [];

  const pid = ss.getSheetByName("Manage").getRange(Number(rowIndex), 1).getValue();
  const fpSheet = ss.getSheetByName("Floorplans");
  if (!fpSheet) return [];
  
  const fpData = fpSheet.getDataRange().getValues();

  return fpData.slice(1)
    .filter(row => String(row[0]) === String(pid))
    .map(row => ({
      id: row[6], // RoomID
      display: row[2] ? `${row[1]} (#${row[2]})` : row[1]
    }));
}

/**
 * MISSING PIECE 1: Handles the error you got.
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
 * MISSING PIECE 2: Required by your existing deleteTask function.
 * Simply performs the row deletion safely.
 */
function safeDeleteRow(sheetName, rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet && rowIdx > 1) {
    sheet.deleteRow(rowIdx);
  }
}


/**
 * Old artifacts
 */
function getTasksData_Off() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    if (!rowIndex) return [];

    const manageSheet = ss.getSheetByName("Manage");
    const selectedPid = manageSheet.getRange(Number(rowIndex), 1).getValue();


    /**naming columns here*/
    const taskSheet = ss.getSheetByName("Tasks");
    const data = taskSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    // Mapping: Col A (PID), B (Task), C (Value), D (RoomName), E (RoomID), F (TaskID)
    return data.slice(1).map((row, index) => {
      return {
        pid: row[0] ? String(row[0]) : "",
        taskType: row[1] ? String(row[1]) : "",
        value: row[2] || 0,
        roomName: row[3] || "No Room Assigned", // Now pulled directly from Col D
        roomId: row[4] ? String(row[4]) : "",    // Col E
        taskId: row[5] ? String(row[5]) : "",    // Col F
        sheetRow: index + 2 
      };
    }).filter(task => String(task.pid) === String(selectedPid));
  } catch (e) {
    console.error("Error in getTasksData: " + e.toString());
    return [];
  }
}

function saveTask_Off(taskObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Tasks");
  
  // 1. Resolve Room Name from Floorplans to ensure we have the latest display name
  const fpSheet = ss.getSheetByName("Floorplans");
  const fpData = fpSheet.getDataRange().getValues();
  const roomRow = fpData.slice(1).find(r => String(r[6]) === String(taskObj.roomId));
  const roomDisplayName = roomRow ? (roomRow[2] ? `${roomRow[1]} (#${roomRow[2]})` : roomRow[1]) : "No Room Assigned";

  // 2. Manage TaskID (Col F)
  let taskId = taskObj.taskId;
  if (!taskId && taskObj.sheetRow) {
    taskId = sheet.getRange(Number(taskObj.sheetRow), 6).getValue();
  }
  if (!taskId) {
    taskId = "T-" + new Date().getTime(); 
  }

  // 3. MAPPING: Task (B), Value (C), RoomName (D), RoomID (E), TaskID (F)
  const rowValues = [[
    String(taskObj.taskType), 
    Number(taskObj.value), 
    String(roomDisplayName), 
    String(taskObj.roomId), 
    String(taskId)
  ]];

  if (taskObj.sheetRow) {
    // UPDATE starting at Col B (2), spanning 5 columns
    sheet.getRange(Number(taskObj.sheetRow), 2, 1, 5).setValues(rowValues);
    
    // 4. CASCADE: If the task changed rooms or names, update associated Materials/Equipment
    // This uses the utility we put in Service_Utils.gs
    cascadeRoomNameUpdate(taskObj.roomId, roomDisplayName);
    
  } else {
    // APPEND
    const userProperties = PropertiesService.getUserProperties();
    const manageRow = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    const pid = ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue();
    sheet.appendRow([pid, ...rowValues[0]]);
  }
  
  return getTasksData();
}

function deleteTask_Off(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName("Tasks");
  const targetRow = Number(rowIdx);
  
  // 1. Safety Check: Does this Task have materials or equipment assigned?
  const taskId = taskSheet.getRange(targetRow, 6).getValue();
  
  // Check Materials
  const matSheet = ss.getSheetByName("Materials");
  const matData = matSheet.getDataRange().getValues();
  const hasMaterials = matData.some(row => String(row[4]).trim() === String(taskId).trim());

  // Check Equipment
  const equipSheet = ss.getSheetByName("Equipment");
  const equipData = equipSheet.getDataRange().getValues();
  const hasEquipment = equipData.some(row => String(row[4]).trim() === String(taskId).trim());

  if (hasMaterials || hasEquipment) {
    return { 
      success: false, 
      error: "Cannot delete task: It has Materials or Equipment assigned to it." 
    };
  }

  // 2. Call Global Helper from Service_Utils.gs
  safeDeleteRow("Tasks", targetRow);
  
  return { success: true };
}