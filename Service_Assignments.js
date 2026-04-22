/**
 * Service_Assignments.gs
 */
function getAssignmentsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Assignments");
    const activePid = getActivePid(); // Uses the logic from Service_Utils
    
    if (!activePid || !sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    // Columns: A:PiD, B:Role, C:First, D:Last, E:Middle, F:Email, G:Phone, H:CompanyCode
    return data.slice(1)
      .map((row, index) => {
        return {
          pid: String(row[0] || ""),
          role: String(row[1] || ""),
          first: String(row[2] || ""),
          last: String(row[3] || ""),
          middle: String(row[4] || ""),
          email: String(row[5] || ""),
          phone: String(row[6] || ""),
          companyCode: String(row[7] || ""),
          sheetRow: index + 2
        };
      })
      .filter(a => a.pid === activePid);
      
  } catch (e) {
    console.error("Error in getAssignmentsData: " + e.toString());
    return [];
  }
}

function saveAssignmentData(assignObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Assignments");
  const activePid = getActivePid();

  const rowValues = [[
    String(assignObj.role), 
    String(assignObj.first), 
    String(assignObj.last), 
    String(assignObj.middle),
    String(assignObj.email),
    String(assignObj.phone),
    String(assignObj.companyCode)
  ]];

  if (assignObj.sheetRow && Number(assignObj.sheetRow) > 1) {
    // Update Col B through H (Indices 2 to 8)
    sheet.getRange(Number(assignObj.sheetRow), 2, 1, 7).setValues(rowValues);
  } else {
    // Append [PiD, Role, First, Last, Middle, Email, Phone, CompanyCode]
    sheet.appendRow([activePid, ...rowValues[0]]);
  }
  return getAssignmentsData();
}

function deleteAssignment(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Assignments");
  sheet.deleteRow(parseInt(rowIdx));
  return getAssignmentsData();
}

/**
 * Fetches companies for the dropdown: "Client Name (CODE)"
 */
function getAssignmentCompanyList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Contacts"); // Assuming contacts holds companies
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  
  // Assuming Col 1 is Name, Col 2 is Code
  return data.slice(1).map(row => ({
    name: `${row[0]} (${row[1]})`,
    code: row[1]
  })).filter(c => c.code);
}