function getInsuranceData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Insurance");
    const manageSheet = ss.getSheetByName("Manage");
    
    const userProperties = PropertiesService.getUserProperties();
    const rowIndex = userProperties.getProperty('ACTIVE_PROJECT_ROW');
    
    if (!rowIndex || !manageSheet || !sheet) return [];

    // Get the PID and force it to a clean string
    const selectedPid = String(manageSheet.getRange(Number(rowIndex), 1).getValue()).trim();
    if (!selectedPid) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const timezone = ss.getSpreadsheetTimeZone();

    return data.slice(1)
      .map((row, index) => {
        // Handle the Date conversion safely
        let dateVal = row[1];
        if (dateVal instanceof Date) {
          dateVal = Utilities.formatDate(dateVal, timezone, "MM/dd/yyyy");
        }

        return {
          pid: row[0] ? String(row[0]).trim() : "",
          dateLoss: dateVal || "", 
          type: row[2] || "", 
          company: row[3] || "", 
          policy: row[4] || "", 
          claim: row[5] || "", 
          firstName: row[7] || "", 
          lastName: row[8] || "", 
          email: row[9] || "", 
          insuranceId: row[10] || "",
          sheetRow: index + 2
        };
      })
      .filter(ins => ins.pid === selectedPid); // Compare clean strings
      
  } catch (e) {
    console.error("Error: " + e.toString());
    return [];
  }
}

/**
 * Saves or updates insurance records
 */
function updateInsuranceData(insObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Insurance");
    
    // Mapping matches your spreadsheet screenshot:
    // Col 1: PID (handled in append), Col 2: Date, Col 3: Type, Col 4: Co, 
    // Col 5: Policy, Col 6: Claim, Col 7: C=I (blank), Col 8: First, 
    // Col 9: Last, Col 10: Email, Col 11: ID
    const rowValues = [
      insObj.dateLoss, 
      insObj.type, 
      insObj.company, 
      insObj.policy, 
      insObj.claim, 
      "", // Column G (Checkbox index 6)
      insObj.firstName, 
      insObj.lastName, 
      insObj.email, 
      insObj.insuranceId
    ];

    if (insObj.sheetRow && Number(insObj.sheetRow) > 1) {
      // Update existing: Start at Col B (index 2), length of 10 columns
      sheet.getRange(Number(insObj.sheetRow), 2, 1, 10).setValues([rowValues]);
    } else {
      // Create new: Get active PID first
      const userProperties = PropertiesService.getUserProperties();
      const manageRow = userProperties.getProperty('ACTIVE_PROJECT_ROW');
      const pid = ss.getSheetByName("Manage").getRange(Number(manageRow), 1).getValue();
      
      sheet.appendRow([pid, ...rowValues]);
    }
    return true; // Success
  } catch (e) {
    throw new Error("Update failed: " + e.toString());
  }
}

/**
 * Deletes an insurance record
 */
function deleteInsuranceRow(row) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Insurance");
    sheet.deleteRow(Number(row));
    return true;
  } catch (e) {
    throw new Error("Delete failed: " + e.toString());
  }
}