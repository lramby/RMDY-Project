/**
 * Service_Site.gs 
 * Refactored to use CONFIG and Standardized Objects
 */

/*=======================================
* Site Functions
*=======================================*/
/**
 * Retrieves Site (Project Overview) data for the active project row.
 */
function getSiteData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    const activePid = userProperties.getProperty('ACTIVE_PROJECT_ID');
    
    if (!rowIndex) return null;

    const sheet = ss.getSheetByName(CONFIG.TABLES.SITES.NAME);
    const COLS = CONFIG.TABLES.SITES.COLUMNS;
    const targetRow = parseInt(rowIndex);

    // Safety: If row is out of bounds, return an empty object with the PID
    if (targetRow > sheet.getLastRow() || targetRow < 2) {
      return {
        pid: activePid || 'N/A',
        sqFt: '', construction: '', occupancy: '',
        yearBuilt: '', usage: '', residence: '', basement: ''
      };
    }

    // Get the whole row based on config length
    const data = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Map using CONFIG.TABLES.SITES.COLUMNS
    return {
      pid:         			data[COLS.PID] || activePid,
      sqFt:         		data[COLS.APPROXAREA],
      constructionType: data[COLS.CONSTRUCTIONTYPE],
      occupancy:    		data[COLS.OCCUPANCY],
      yearBuilt:    		data[COLS.YEARBUILT],
      usageType:        data[COLS.USAGETYPE],
      residenceType:    data[COLS.RESIDENCETYPE],
      basement:     		data[COLS.BASEMENT]
    };
  } catch (e) {
    console.error("Error in getSiteData: " + e.toString());
    return null;
  }
}

/**
 * Updates Site data based on current project row.
 */
function updateSiteData(formData) {
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  const COLS = CONFIG.TABLES.SITES.COLUMNS;
  
  let activePid = userProperties.getProperty('ACTIVE_PROJECT_ID');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const siteSheet = ss.getSheetByName(CONFIG.TABLES.SITES.NAME);
  const targetRow = parseInt(rowIndex);

  if (!activePid && rowIndex) {
    activePid = siteSheet.getRange(targetRow, 1).getValue();
  }
  
  if (!activePid || !rowIndex) {
    throw new Error("Save Blocked: Missing Project ID reference.");
  }

  // Map data into the row array
  const rowArray = [];
  rowArray[COLS.PID] = activePid;
  rowArray[COLS.APPROXAREA] = formData.approxArea;
  rowArray[COLS.CONSTRUCTIONTYPE] = formData.constructionType;
  rowArray[COLS.OCCUPANCY] = formData.occupancy;
  rowArray[COLS.YEARBUILT] = formData.yearBuilt;
  rowArray[COLS.USAGETYPE] = formData.usageType;
  rowArray[COLS.RESIDENCETYPE] = formData.residenceType;
  rowArray[COLS.BASEMENT] = formData.basement;

  // Use the helper to save
  saveCommonData(CONFIG.TABLES.SITES.NAME, targetRow, rowArray);
  
  return getSiteData();
}




/*=======================================
* Rooms Functions
*=======================================*/
/**
 * Retrieves all rooms associated with the active Project ID.
 */
function getRoomData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.TABLES.ROOMS.NAME);
    const COLS = CONFIG.TABLES.ROOMS.COLUMNS;
    const props = PropertiesService.getUserProperties();
    
    let activePid = props.getProperty('ACTIVE_PROJECT_ID');
    if (!activePid) {
      const siteRow = props.getProperty('ACTIVE_PROJECT_ROW');
      activePid = ss.getSheetByName(CONFIG.TABLES.SITES.NAME).getRange(siteRow, 1).getValue();
    }

    if (!sheet || !activePid) return [];
    
    const data = sheet.getDataRange().getValues();
    
    return data.slice(1).map((row, idx) => {
      // Filter by PID
      if (String(row[COLS.PID]).trim() === String(activePid).trim()) {
        return {
          sheetRow: idx + 2,
          pid: row[COLS.PID],
          roomName: row[COLS.ROOMNAME],
          roomNumber: row[COLS.ROOMNUMBER],
          length: row[COLS.LENGTH],
          width: row[COLS.WIDTH],
          height: row[COLS.HEIGHT],
          roomId: row[COLS.ROOMID],
          lengthUnit: row[COLS.LENGTHUNIT] || "ft",
          widthUnit: row[COLS.WIDTHUNIT] || "ft",
          heightUnit: row[COLS.HEIGHTUNIT] || "ft"
        };
      }
      return null;
    }).filter(r => r !== null);
  } catch (e) {
    console.error("Error in getRoomData: " + e.toString());
    return [];
  }
}

/**
 * Saves or updates a room and triggers the name cascade.
 */
function saveRoomData(roomObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABLES.ROOMS.NAME);
  const COLS = CONFIG.TABLES.ROOMS.COLUMNS;
  const props = PropertiesService.getUserProperties();
  
  let pid = props.getProperty('ACTIVE_PROJECT_ID');
  if (!pid) {
    const siteRow = props.getProperty('ACTIVE_PROJECT_ROW');
    pid = ss.getSheetByName(CONFIG.TABLES.SITES.NAME).getRange(siteRow, 1).getValue();
  }

  // 1. Manage RoomID (Create if new)
  let roomId = roomObj.sheetRow ? roomObj.roomId : (pid + "-" + new Date().getTime());
  
  // Create the display name for cascading (e.g., "Kitchen (#101)")
  const newDisplayName = roomObj.roomNumber ? `${roomObj.roomName} (#${roomObj.roomNumber})` : roomObj.roomName;

  // 2. Map object to row array using CONFIG
  const rowArray = [];
  rowArray[COLS.PID] = pid;
  rowArray[COLS.ROOMNAME] = roomObj.roomName;
  rowArray[COLS.ROOMNUMBER] = roomObj.roomNumber;
  rowArray[COLS.LENGTH] = roomObj.length;
  rowArray[COLS.WIDTH] = roomObj.width;
  rowArray[COLS.HEIGHT] = roomObj.height;
  rowArray[COLS.ROOMID] = roomId;
  rowArray[COLS.LENGTHUNIT] = roomObj.lengthUnit || "ft";
  rowArray[COLS.WIDTHUNIT] = roomObj.widthUnit || "ft";
  rowArray[COLS.HEIGHTUNIT] = roomObj.heightUnit || "ft";

  if (roomObj.sheetRow && Number(roomObj.sheetRow) >= 2) {
    // UPDATE
    sheet.getRange(Number(roomObj.sheetRow), 1, 1, rowArray.length).setValues([rowArray]);
    // 3. THE CENTRAL CASCADE: Update Tasks, Materials, and Equipment
    cascadeRoomNameUpdate(roomId, newDisplayName);
  } else {
    // NEW
    sheet.appendRow(rowArray);
  }
  
  return getRoomData();
}

/**
 * Safely deletes a room only if no tasks are assigned.
 */
function deleteRoom(rowIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const roomSheet = ss.getSheetByName(CONFIG.TABLES.ROOMS.NAME);
    const taskSheet = ss.getSheetByName(CONFIG.TABLES.TASKS.NAME);
    
    const ROOM_COLS = CONFIG.TABLES.ROOMS.COLUMNS;
    const TASK_COLS = CONFIG.TABLES.TASKS.COLUMNS;

    // 1. Get the Unique RoomID from the sheet
    const roomIdToDelete = roomSheet.getRange(rowIdx, ROOM_COLS.ROOMID + 1).getValue();

    // 2. Check if any Tasks are using this RoomID
    const taskData = taskSheet.getDataRange().getValues();
    const hasTasks = taskData.some(row => String(row[TASK_COLS.ROOMID]).trim() === String(roomIdToDelete).trim());

    if (hasTasks) {
      return { 
        success: false, 
        error: "Cannot delete a room with tasks assigned to it!" 
      };
    }

    roomSheet.deleteRow(rowIdx);
    return { success: true };

  } catch (e) {
    return { success: false, error: "System Error: " + e.toString() };
  }
}