/**
 * Service_Site.gs 
 * Focuses strictly on Filtering (Read) and Preparation (Write)
 * Mapping logic is handled by Code.gs and CONFIG
 */

/*=======================================
 * Site Functions
 *=======================================*/

function getActiveSiteData() {
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = parseInt(userProperties.getProperty('ACTIVE_PROJECT_ROW'));
  
  if (!rowIndex) return null;

  // Uses the centralized mapping from Code.gs
  const allSites = getSiteData(); 
  
  // Just find and return the specific project row
  return allSites.find(site => site.rowIndex === rowIndex) || null;
}

function updateSiteData(formData) {
  const userProperties = PropertiesService.getUserProperties();
  const rowIndex = parseInt(userProperties.getProperty('ACTIVE_PROJECT_ROW'));
  const activePid = userProperties.getProperty('ACTIVE_PROJECT_ID');
  const COLS = CONFIG.TABLES.SITES.COLUMNS;

  if (!rowIndex || !activePid) throw new Error("Missing Project Reference");

  // Prepare flat array for the spreadsheet using CONFIG indices
  const rowArray = [];
  rowArray[COLS.PID] = activePid;
  rowArray[COLS.APPROXAREA] = formData.approxArea;
  rowArray[COLS.CONSTRUCTIONTYPE] = formData.constructionType;
  rowArray[COLS.OCCUPANCYTYPE] = formData.occupancyType;
  rowArray[COLS.YEARBUILT] = formData.yearBuilt;
  rowArray[COLS.USAGETYPE] = formData.usageType;
  rowArray[COLS.RESIDENCETYPE] = formData.residenceType;
  rowArray[COLS.BASEMENTTYPE] = formData.basementType;

  saveCommonData(CONFIG.TABLES.SITES.NAME, rowIndex, rowArray);
  
  return getActiveSiteData();
}

/*=======================================
 * Rooms Functions
 *=======================================*/

function getActiveRoomData() {
  const activePid = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ID');
  if (!activePid) return [];

  // Uses the centralized mapping from Code.gs
  const allRooms = getRoomsData();

  // Just filter for the current project
  return allRooms.filter(room => String(room.pid).trim() === String(activePid).trim());
}

function saveRoomData(roomObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TABLES.ROOMS.NAME);
  const COLS = CONFIG.TABLES.ROOMS.COLUMNS;
  const activePid = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ID');

  let roomId = roomObj.roomID || (activePid + "-" + new Date().getTime());
  const newDisplayName = roomObj.roomNumber ? `${roomObj.roomName} (#${roomObj.roomNumber})` : roomObj.roomName;

  // Prepare flat array for spreadsheet
  const rowArray = [];
  rowArray[COLS.PID] = activePid;
  rowArray[COLS.ROOMNAME] = roomObj.roomName;
  rowArray[COLS.ROOMNUMBER] = roomObj.roomNumber;
  rowArray[COLS.LENGTH] = roomObj.length;
  rowArray[COLS.WIDTH] = roomObj.width;
  rowArray[COLS.HEIGHT] = roomObj.height;
  rowArray[COLS.ROOMID] = roomId;
  rowArray[COLS.LENGTHUNIT] = roomObj.lengthUnit;
  rowArray[COLS.WIDTHUNIT] = roomObj.widthUnit;
  rowArray[COLS.HEIGHTUNIT] = roomObj.heightUnit;

  if (roomObj.rowIndex && Number(roomObj.rowIndex) >= 2) {
    sheet.getRange(Number(roomObj.rowIndex), 1, 1, rowArray.length).setValues([rowArray]);
    if (typeof cascadeRoomNameUpdate === 'function') cascadeRoomNameUpdate(roomId, newDisplayName);
  } else {
    sheet.appendRow(rowArray);
  }
  
  return getActiveRoomData();
}