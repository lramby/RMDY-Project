/**
 * Service_Site.gs
 */
/**
 * Service_Site.gs
 * Safely fetches site-specific data using the active project's row index.
 */
function getSiteData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Site");
  const userProperties = PropertiesService.getUserProperties();
  
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  const activePid = userProperties.getProperty('ACTIVE_PROJECT_ID');
  
  if (!rowIndex) return null;

  const targetRow = parseInt(rowIndex);
  const lastRow = sheet.getLastRow();

  // --- SAFETY CHECK ---
  if (targetRow > lastRow || targetRow < 2) {
    return {
      pid: activePid || 'N/A',
      sqFt: '', construction: '', occupancy: '',
      yearBuilt: '', usage: '', residence: '', basement: ''
    };
  }

  // Adjusted to match your Column A-H layout
  const displayValues = sheet.getRange(targetRow, 1, 1, 8).getDisplayValues()[0];
  
  return {
    pid:          displayValues[0] || activePid,
    sqFt:         displayValues[1],
    construction: displayValues[2],
    occupancy:    displayValues[3],
    yearBuilt:    displayValues[4],
    usage:        displayValues[5],
    residence:    displayValues[6],
    basement:     displayValues[7]
  };
}

function getRoomData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Floorplans");
  const props = PropertiesService.getUserProperties();
  let activePid = props.getProperty('ACTIVE_PROJECT_ID');
  
  if (!activePid) {
    const siteRow = props.getProperty('ACTIVE_PROJECT_ROW');
    activePid = ss.getSheetByName("Site").getRange(siteRow, 1).getValue();
  }

  if (!sheet || !activePid) return [];
  
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map((row, idx) => {
    if (String(row[0]).trim() === String(activePid).trim()) {
      return {
        sheetRow: idx + 2,
        name: row[1],
        roomNum: row[2],
        l: row[3],
        w: row[4],
        h: row[5],
        roomId: row[6] // Essential for the name update logic
      };
    }
    return null;
  }).filter(r => r !== null);
}

function saveRoomData(room) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fpSheet = ss.getSheetByName("Floorplans");
  const props = PropertiesService.getUserProperties();
  
  let pid = props.getProperty('ACTIVE_PROJECT_ID');
  if (!pid) pid = ss.getSheetByName("Site").getRange(props.getProperty('ACTIVE_PROJECT_ROW'), 1).getValue();

  // 1. Manage RoomID
  let roomId = room.row ? fpSheet.getRange(room.row, 7).getValue() : (pid + "-" + new Date().getTime());
  const newDisplayName = room.roomNum ? `${room.name} (#${room.roomNum})` : room.name;

  // 2. Prep values for Floorplans (7 columns)
  const vals = [[pid, room.name, room.roomNum, room.l, room.w, room.h, roomId]];
  
  if (room.row) {
    // UPDATE Floorplan
    fpSheet.getRange(room.row, 1, 1, 7).setValues(vals);

    // 3. THE CENTRAL CASCADE: Update Tasks, Materials, and Equipment via Service_Utils
    cascadeRoomNameUpdate(roomId, newDisplayName);
    
  } else {
    // APPEND New Room
    fpSheet.appendRow(vals[0]);
  }
  return getRoomData();
}

/** * SAFETY DELETE: Checks for assigned tasks before removing the room.
 */
function deleteRoom(rowIdx) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fpSheet = ss.getSheetByName("Floorplans");
    const taskSheet = ss.getSheetByName("Tasks");

    // 1. Get the RoomID from Column G (7) of the Floorplans sheet
    const roomIdToDelete = fpSheet.getRange(rowIdx, 7).getValue();

    // 2. Check if RoomID exists in the Tasks sheet (Column E / Index 4)
    const taskData = taskSheet.getDataRange().getValues();
    const hasTasks = taskData.some(row => String(row[4]).trim() === String(roomIdToDelete).trim());

    if (hasTasks) {
      // Return a custom message to the front end
      return { 
        success: false, 
        error: "Cannot delete a room with tasks assigned to it!" 
      };
    }

    // 3. If no tasks, proceed with deletion
    fpSheet.deleteRow(rowIdx);
    return { success: true };

  } catch (e) {
    return { success: false, error: "System Error: " + e.toString() };
  }
}

/**
 * Service_Site.gs 
 * Updates the Site data on the spreadsheet based on current project row.
 */
function updateSiteData(formData) {
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
  
  // 1. Try to get PID from properties
  let activePid = userProperties.getProperty('ACTIVE_PROJECT_ID');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const siteSheet = ss.getSheetByName("Site");
  const targetRow = parseInt(rowIndex);

  // 2. EMERGENCY BACKUP: If properties failed, grab it from the sheet row directly
  if (!activePid && rowIndex) {
    activePid = siteSheet.getRange(targetRow, 1).getValue();
  }
  
  if (!activePid || !rowIndex) {
    throw new Error("Save Blocked: Missing Project ID or Row Reference.");
  }

  // 3. Prepare the row data
  const rowData = [
    activePid, // This is now guaranteed to have a value
    formData.sqFt,
    formData.construction,
    formData.occupancy,
    formData.yearBuilt,
    formData.usage,
    formData.residence,
    formData.basement
  ];

  // 4. Save using your Universal Saver
  saveCommonData("Site", targetRow, rowData);
  
  return getSiteData();
}