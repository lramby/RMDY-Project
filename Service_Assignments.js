/**
 * Service_Assignments.gs
 * Refactored to use CONFIG object and Helper functions.
 */

/**
 * Gets assignemnt data and converts it from array to object
 */
function getAssignmentsDataForActivePid() {
  try {
    const activePid = getActivePid();
    console.log("DEBUG - getAssignmentsDataForActivePid - ActivePid:", activePid);
    
    if (!activePid) return [];

    // 1. FETCH & FILTER (using your Universal Utility)
    // We assume PID is in Column A (index 0)
    const rawMatches = getFilteredData(CONFIG.TABLES.ASSIGNMENTS.NAME, 0); 

    // 2. MAP (Transforming the data for the UI)
    const cols = CONFIG.TABLES.ASSIGNMENTS.COLUMNS;
    
    return rawMatches.map(obj => {
      const row = obj.data;
      return {
        rowIndex: obj.rowIdx,
        pid: row[cols.PID],
        roleName: row[cols.ROLENAME],
        firstName: row[cols.FIRSTNAME],
        middleName: row[cols.MIDDLENAME],
        lastName: row[cols.LASTNAME],
        email: row[cols.EMAIL],
        phone: row[cols.PHONE],
        companyCode: row[cols.COMPANYCODE]
      };
    });

  } catch (e) {
    console.error("Error in getAssignmentsDataForActivePid: " + e.toString());
    return [];
  }
}

/**
 * Saves or Updates an assignment
 * @param {Object} assignObj - The assignment data from the UI
 */
function saveAssignmentData(assignObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABLES.ASSIGNMENTS.NAME);
  const cols = CONFIG.TABLES.ASSIGNMENTS.COLUMNS;
  const activePid = getActivePid();
  const clean = (val) => String(val || "").trim();

  // processSaveAssignment sends "sheetRow"
  const rowNum = Number(assignObj.sheetRow);

  if (rowNum && rowNum > 1) {
    // UPDATE EXISTING
    sheet.getRange(rowNum, cols.ROLENAME + 1).setValue(clean(assignObj.roleName));
    sheet.getRange(rowNum, cols.FIRSTNAME + 1).setValue(clean(assignObj.firstName));
    sheet.getRange(rowNum, cols.LASTNAME + 1).setValue(clean(assignObj.lastName));
    sheet.getRange(rowNum, cols.MIDDLENAME + 1).setValue(clean(assignObj.middleName));
    sheet.getRange(rowNum, cols.EMAIL + 1).setValue(clean(assignObj.email));
    sheet.getRange(rowNum, cols.PHONE + 1).setValue(clean(assignObj.phone));
    sheet.getRange(rowNum, cols.COMPANYCODE + 1).setValue(clean(assignObj.companyCode));
  } else {
    // CREATE NEW
    const timestamp = new Date().getTime().toString().slice(-5); 
    const newId = `${activePid}-ASGN-${timestamp}`;

    const newRow = [];
    newRow[cols.PID] = activePid;
    newRow[cols.ASSIGNMENTID] = newId;
    newRow[cols.ROLENAME] = clean(assignObj.roleName);
    newRow[cols.FIRSTNAME] = clean(assignObj.firstName);
    newRow[cols.LASTNAME] = clean(assignObj.lastName);
    newRow[cols.MIDDLENAME] = clean(assignObj.middleName);
    newRow[cols.EMAIL] = clean(assignObj.email);
    newRow[cols.PHONE] = clean(assignObj.phone);
    newRow[cols.COMPANYCODE] = clean(assignObj.companyCode);

    sheet.appendRow(newRow);
  }
  return getAssignmentsDataForActivePid();
}

/**
 * Deletes an assignment row
 */
function deleteAssignment(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABLES.ASSIGNMENTS.NAME);
  sheet.deleteRow(parseInt(rowIdx));
  return getAssignmentsDataForActivePid();
}

/**
 * Fetches companies for the dropdown using CONTACTS config
 */
function getAssignmentCompanyList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TABLES.CONTACTS.NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const cols = CONFIG.TABLES.CONTACTS.COLUMNS;
  
  return data.slice(1).map(row => {
    const companyName = row[cols.COMPANYNAME];
    const companyCode = row[cols.COMPANYCODE];
    return {
      companyName: `${companyName} (${companyCode})`,
      companyCode: companyCode
    };
  }).filter(c => c.companyCode);
}

function loadAssignments() {
  // ... loader logic ...
  google.script.run
    .withSuccessHandler(renderAssignments)
    .getAssignmentsDataForActivePid(); // Call the FILTERED version
}