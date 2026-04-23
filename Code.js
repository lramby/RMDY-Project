/**
 * Code.gs
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gana Consulting - Pivot')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPageHtml(pageName) {
  return HtmlService.createHtmlOutputFromFile('Page_' + pageName).getContent();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** gets the active project*/
function getProjectList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manage");
  const data = sheet.getDataRange().getValues();
  
  // Get the stored index from memory
  const activeRow = PropertiesService.getUserProperties().getProperty('ACTIVE_PROJECT_ROW');

  const projects = data.slice(1).map((row, index) => {
    return {
      pid: row[0],
      client: row[1],
      address: row[2],
      sheetRow: index + 2
    };
  });

  return {
    projects: projects,
    activeRow: activeRow // Send this back to the UI
  };
}

function getTasksData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TABLES.TASKS.NAME);
  const values = sheet.getDataRange().getValues();
  const cols = CONFIG.TABLES.TASKS.COLUMNS;
  
  // Skip header and map to objects
  return values.slice(1).map((row, index) => {
    return {
      rowIndex: index + 2, // 1 for header, 1 for 0-indexing
      pid: row[cols.PID],
      task: row[cols.TASK],
      value: row[cols.VALUE],
      roomName: row[cols.ROOMNAME],
      roomID: row[cols.ROOMID],
      taskID: row[cols.TASKID],
      unit: row[cols.UNIT],
      price: row[cols.PRICE],
      cost: row[cols.COST],
      note: row[cols.NOTE],
    };
  });
}

function getEquipmentData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TABLES.EQUIPMENT.NAME);
  const values = sheet.getDataRange().getValues();
  const cols = CONFIG.TABLES.EQUIPMENT.COLUMNS;
  
  // Skip header and map to objects
  return values.slice(1).map((row, index) => {
    return {
      rowIndex: index + 2, // 1 for header, 1 for 0-indexing
      pid: row[cols.PID],
      itemID: row[cols.ITEMID],
      roomID: row[cols.ROOMID],
      taskID: row[cols.TASKID],
      roomName: row[cols.ROOMNAME],
      taskName: row[cols.TASKNAME],
      item: row[cols.ITEM],
      value: row[cols.VALUE],
      unit: row[cols.UNIT],
      price: row[cols.PRICE],
      cost: row[cols.COST],
      note: row[cols.NOTE],
    };
  });
}

function getDetailsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TABLES.DETAILS.NAME);
  const values = sheet.getDataRange().getValues();
  const cols = CONFIG.TABLES.DETAILS.COLUMNS;
  
  // Skip header and map to objects
  return values.slice(1).map((row, index) => {
    return {
      rowIndex: index + 2, // 1 for header, 1 for 0-indexing
      pid: row[cols.PID],
      address1: row[cols.ADDRESS1],
      address2: row[cols.ADDRESS2],
      city: row[cols.CITY],
      state: row[cols.STATE],
      zip: row[cols.ZIP],
      country: row[cols.COUNTRY],
      firstName: row[cols.FIRSTNAME],
      lastName: row[cols.LASTNAME],
      email: row[cols.EMAIL],
      phone: row[cols.PHONE]
    };
  });
}

function getDatesData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TABLES.DATES.NAME);
  const values = sheet.getDataRange().getValues();
  const cols = CONFIG.TABLES.DATES.COLUMNS;
  
  // Skip header and map to objects
  return values.slice(1).map((row, index) => {
    return {
      rowIndex: index + 2, // 1 for header, 1 for 0-indexing
      pid: row[cols.PID],
      loss: row[cols.LOSS],
      duedate: row[cols.DUEDATE],
      contacted: row[cols.CONTACTED],
      assigned: row[cols.ASSIGNED],
      inspected: row[cols.INSPECTED],
      estimated: row[cols.ESTIMATED],
      started: row[cols.STARTED],
      finished: row[cols.FINISHED],
      invoiced: row[cols.INVOICED],
      approved: row[cols.APPROVED],
			paid: row[cols.PAID]
    };
  });
}