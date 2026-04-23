/**
 * Service_Assignments.gs
 * Refactored to use CONFIG object and Helper functions.
 */

function getAssignmentsDataForActivePid() {
  try {
    const activePid = getActivePid(); // Logic from Service_Utils
    if (!activePid) return [];

    // Use the helper function from Code.gs to get structured data
    const allAssignments = getAssignmentsData(); 
    
    // Filter for the active PID
    return allAssignments.filter(a => String(a.pid) === String(activePid));
      
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
  
  // Clean helper to ensure strings and handle nulls
  const clean = (val) => String(val || "").trim();

  if (assignObj.rowIndex && Number(assignObj.rowIndex) > 1) {
    /**
     * UPDATE EXISTING
     */
    const rowNum = Number(assignObj.rowIndex);
    
    // We do not overwrite the PID or ASSIGNMENTID on updates to maintain integrity
    sheet.getRange(rowNum, cols.ROLENAME + 1).setValue(clean(assignObj.roleName));
    sheet.getRange(rowNum, cols.FIRSTNAME + 1).setValue(clean(assignObj.firstName));
    sheet.getRange(rowNum, cols.LASTNAME + 1).setValue(clean(assignObj.lastName));
    sheet.getRange(rowNum, cols.MIDDLENAME + 1).setValue(clean(assignObj.middleName));
    sheet.getRange(rowNum, cols.EMAIL + 1).setValue(clean(assignObj.email));
    sheet.getRange(rowNum, cols.PHONE + 1).setValue(clean(assignObj.phone));
    sheet.getRange(rowNum, cols.COMPANYCODE + 1).setValue(clean(assignObj.companyCode));
    
  } else {
    /**
     * CREATE NEW
     */
    // Generate ID: [PiD]-ASGN-[Timestamp] 
    // Format: "12345-ASGN-88291"
    const timestamp = new Date().getTime().toString().slice(-5); 
    const newId = `${activePid}-ASGN-${timestamp}`;

    // Create a full row array based on CONFIG column indices
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
    const name = row[cols.COMPANYNAME];
    const code = row[cols.COMPANYCODE];
    return {
      name: `${companyName} (${companyCode})`,
      code: companyCode
    };
  }).filter(c => c.code);
}