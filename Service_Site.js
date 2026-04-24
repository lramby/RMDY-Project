/**
 * Service_Site.gs 
 * Aligned with CONFIG.TABLES.SITE (singular)
 */

/*=======================================
 * Site Functions
 *=======================================*/

/**
 * Retrieves the specific site data for the active project row.
 */
function getActiveSiteData() {
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = parseInt(userProperties.getProperty('ACTIVE_PROJECT_ROW'));
  
  if (!rowIndex) return null;

  // Use the mapping logic already in Code.gs
  // Uses CONFIG.TABLES.SITE (singular) internally
  const allSites = getSiteData(); 
  
  // Find the site that matches our active project row
  return allSites.find(site => site.rowIndex === rowIndex) || null;
}

/**
 * Updates Site data using the form data from the UI.
 */
function updateSiteData(formData) {
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = parseInt(userProperties.getProperty('ACTIVE_PROJECT_ROW'));
  const activePid = userProperties.getProperty('ACTIVE_PROJECT_ID');
  
  // Alignment: Using singular SITE key from Config
  const COLS = CONFIG.TABLES.SITE.COLUMNS;

  if (!rowIndex || !activePid) throw new Error("Missing Project Reference");

  // Prepare flat array for the spreadsheet using CONFIG indices
  const rowArray = [];
  rowArray[COLS.PID] = activePid;
  rowArray[COLS.APPROXAREA] = formData.approxArea;
  rowArray[COLS.CONSTRUCTIONTYPE] = formData.constructionType;
  rowArray[COLS.OCCUPANCY] = formData.occupancyType;
  rowArray[COLS.YEARBUILT] = formData.yearBuilt;
  rowArray[COLS.USAGETYPE] = formData.usageType;
  rowArray[COLS.RESIDENCETYPE] = formData.residenceType;
  rowArray[COLS.BASEMENT] = formData.basementType;

  // Save via the universal helper using the singular SITE name
  saveCommonData(CONFIG.TABLES.SITE.NAME, rowIndex, rowArray);
  
  return getActiveSiteData();
}

/*=======================================
 * Rooms Functions
 *=======================================*/

/**
 * Retrieves all rooms filtered by the active Project ID.
 */
function getActiveRoomData() {
  const activePid = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ID');
  if (!activePid) return [];

  // Use the mapping logic already in Code.gs
  const allRooms = getRoomsData();

  // Filter to only show rooms for this project
  return allRooms.filter(room => String(room.pid).trim() === String(activePid).trim());
}

/**
 * Saves/Appends a room and triggers the name cascade.
 */
function saveRoomData(roomObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TABLES.ROOMS.NAME);
  const COLS = CONFIG.TABLES.ROOMS.COLUMNS;
  const activePid = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ID');

  // 1. Manage RoomID
  let roomId = roomObj.roomID || (activePid + "-" + new Date().getTime());
  const newDisplayName = roomObj.roomNumber ? `${roomObj.roomName} (#${roomObj.roomNumber})` : roomObj.roomName;

  // 2. Map object to row array using CONFIG
  const rowArray = [];
  rowArray[COLS.PID] = activePid;
  rowArray[COLS.ROOMNAME] = roomObj.roomName;
  rowArray[COLS.ROOMNUMBER] = roomObj.roomNumber;
  rowArray[COLS.LENGTH] = roomObj.length;
  rowArray[COLS.WIDTH] = roomObj.width;
  rowArray[COLS.HEIGHT] = roomObj.height;
  rowArray[COLS.ROOMID] = roomId;
  rowArray[COLS.LENGTHUNIT] = roomObj.lengthUnit || "ft";
  rowArray[COLS.WIDTHUNIT] = roomObj.widthUnit || "ft";
  rowArray[COLS.HEIGHTUNIT] = roomObj.heightUnit || "ft";

  if (roomObj.rowIndex && Number(roomObj.rowIndex) >= 2) {
    sheet.getRange(Number(roomObj.rowIndex), 1, 1, rowArray.length).setValues([rowArray]);
    if (typeof cascadeRoomNameUpdate === 'function') {
       cascadeRoomNameUpdate(roomId, newDisplayName);
    }
  } else {
    sheet.appendRow(rowArray);
  }
  
  return getActiveRoomData();
}